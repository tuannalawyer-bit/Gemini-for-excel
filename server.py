#!/usr/bin/env python3
"""
Gemini for Excel - Local HTTPS Dev Server
Phục vụ các tệp Add-in qua giao thức HTTPS chuẩn cho Microsoft Excel
"""

import http.server
import ssl
import os
import sys
import datetime
import ipaddress
import io

# Đảm bảo in tiếng Việt trên console Windows không bị lỗi mã hóa
if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

def ensure_ssl_certificates(certs_dir):
    os.makedirs(certs_dir, exist_ok=True)
    cert_file = os.path.join(certs_dir, "cert.pem")
    key_file = os.path.join(certs_dir, "key.pem")

    if os.path.exists(cert_file) and os.path.exists(key_file):
        return cert_file, key_file

    print("[*] Đang tự động tạo chứng chỉ SSL nội bộ cho localhost...")
    try:
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.hazmat.primitives import serialization

        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        subject = issuer = x509.Name([
            x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Gemini for Excel Dev"),
        ])

        now = datetime.datetime.now(datetime.timezone.utc)
        cert = (
            x509.CertificateBuilder()
            .subject_name(subject)
            .issuer_name(issuer)
            .public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(now)
            .not_valid_after(now + datetime.timedelta(days=730))
            .add_extension(
                x509.SubjectAlternativeName([
                    x509.DNSName("localhost"),
                    x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
                ]),
                critical=False,
            )
            .sign(key, hashes.SHA256())
        )

        with open(key_file, "wb") as f:
            f.write(
                key.private_bytes(
                    encoding=serialization.Encoding.PEM,
                    format=serialization.PrivateFormat.TraditionalOpenSSL,
                    encryption_algorithm=serialization.NoEncryption(),
                )
            )

        with open(cert_file, "wb") as f:
            f.write(cert.public_bytes(serialization.Encoding.PEM))

        print(f"[✓] Đã tạo chứng chỉ SSL thành công tại: {certs_dir}")
        return cert_file, key_file

    except Exception as e:
        print(f"[!] Lỗi khi sinh SSL: {e}")
        sys.exit(1)

class CORSRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Bật CORS đầy đủ cho Office Add-in
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200, "OK")
        self.end_headers()

def run_server(port=3000):
    project_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(project_dir)

    certs_dir = os.path.join(project_dir, "certs")
    cert_file, key_file = ensure_ssl_certificates(certs_dir)

    server_address = ("localhost", port)
    httpd = http.server.HTTPServer(server_address, CORSRequestHandler)

    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=cert_file, keyfile=key_file)
    httpd.socket = context.wrap_socket(httpd.socket, server_side=True)

    print("\n" + "=" * 65)
    print("      🚀 GEMINI FOR EXCEL - MÁY CHỦ NỘI BỘ ĐANG HOẠT ĐỘNG")
    print("=" * 65)
    print(f" [+] Đường dẫn máy chủ : https://localhost:{port}/")
    print(f" [+] Thư mục gốc      : {project_dir}")
    print(f" [+] Taskpane URL     : https://localhost:{port}/src/taskpane/taskpane.html")
    print("=" * 65)
    print("\n👉 CÁCH NẠP ADD-IN VÀO EXCEL:")
    print("  1. Mở Microsoft Excel (Desktop hoặc Web).")
    print("  2. Vào thẻ 'Insert' (Chèn) -> 'Get Add-ins' (Tải bổ trợ) -> 'My Add-ins'.")
    print("  3. Chọn 'Upload My Add-in' -> Duyệt chọn tệp 'manifest.xml' trong thư mục này.")
    print("  4. Thẻ 'Gemini AI' sẽ xuất hiện trên thanh công cụ của Excel!\n")
    print("  (Nhấn Ctrl+C để dừng máy chủ bất kỳ lúc nào)\n")

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[*] Đã dừng máy chủ an toàn.")
        httpd.server_close()

if __name__ == "__main__":
    port = 3000
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
    run_server(port)
