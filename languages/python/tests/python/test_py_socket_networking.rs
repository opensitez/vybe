use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: socket + networking — socket, urllib.parse, urllib.request, http.client
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_socket_gethostbyname_and_hostname() {
    let src = r#"
import socket

hostname = socket.gethostname()
print(isinstance(hostname, str))
print(len(hostname) > 0)
localhost_ip = socket.gethostbyname("localhost")
print(localhost_ip in ("127.0.0.1", "::1"))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_socket_inet_aton_ntoa() {
    let src = r#"
import socket

ip_str = "192.168.1.1"
packed = socket.inet_aton(ip_str)
print(len(packed))  # 4 bytes
unpacked = socket.inet_ntoa(packed)
print(unpacked)
"#;
    assert_eq!(run_python(src), vec!["4", "192.168.1.1"]);
}

#[test]
fn test_py_socket_tcp_loopback_echo() {
    let src = r#"
import socket, threading

server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server.bind(("127.0.0.1", 0))
port = server.getsockname()[1]
server.listen(1)

def handle_client():
    conn, _ = server.accept()
    data = conn.recv(1024)
    conn.sendall(b"ECHO: " + data)
    conn.close()

t = threading.Thread(target=handle_client)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
client.connect(("127.0.0.1", port))
client.sendall(b"Hello Server")
response = client.recv(1024)
client.close()
t.join()
server.close()

print(response.decode())
"#;
    assert_eq!(run_python(src), vec!["ECHO: Hello Server"]);
}

#[test]
fn test_py_urllib_parse_urlsplit_urlunsplit() {
    let src = r#"
from urllib.parse import urlsplit, urlunsplit, parse_qs

url = "https://user:pass@example.com:8080/path/to/page?query=python&category=code#section"
split = urlsplit(url)

print(split.scheme)
print(split.netloc)
print(split.path)
print(split.query)
print(split.fragment)

qs = parse_qs(split.query)
print(qs["query"])

reconstructed = urlunsplit(split)
print(reconstructed == url)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "https",
            "user:pass@example.com:8080",
            "/path/to/page",
            "query=python&category=code",
            "section",
            "['python']",
            "True"
        ]
    );
}

#[test]
fn test_py_urllib_parse_urlencode_quote() {
    let src = r#"
from urllib.parse import urlencode, quote, unquote

params = {"name": "Alice & Bob", "city": "New York"}
encoded = urlencode(params)
print(encoded)

raw = "hello world / python & code"
q = quote(raw)
print(q)
print(unquote(q) == raw)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "name=Alice+%26+Bob&city=New+York",
            "hello%20world%20/%20python%20%26%20code",
            "True"
        ]
    );
}

#[test]
fn test_py_http_client_response_parsing() {
    let src = r#"
from http.client import HTTPResponse, parse_headers
import io

header_bytes = b"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: 18\r\n\r\n{\"status\": \"ok\"}"
headers = parse_headers(io.BytesIO(header_bytes.split(b"\r\n\r\n")[0] + b"\r\n\r\n"))
print(headers["Content-Type"])
print(headers["Content-Length"])
"#;
    assert_eq!(run_python(src), vec!["application/json", "18"]);
}

#[test]
fn test_py_socket_timeout() {
    let src = r#"
import socket

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.settimeout(0.01)
print(s.gettimeout())
s.close()
"#;
    assert_eq!(run_python(src), vec!["0.01"]);
}

#[test]
fn test_py_socket_udp_loopback() {
    let src = r#"
import socket, threading

server = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
server.bind(("127.0.0.1", 0))
port = server.getsockname()[1]

def udp_server():
    data, addr = server.recvfrom(1024)
    server.sendto(b"ACK: " + data, addr)

t = threading.Thread(target=udp_server)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
client.sendto(b"UDP Ping", ("127.0.0.1", port))
data, _ = client.recvfrom(1024)
client.close()
t.join()
server.close()

print(data.decode())
"#;
    assert_eq!(run_python(src), vec!["ACK: UDP Ping"]);
}

#[test]
fn test_py_socket_options_reuseaddr() {
    let src = r#"
import socket

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
opt = s.getsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR)
print(opt != 0)
s.close()
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_urllib_parse_urljoin() {
    let src = r#"
from urllib.parse import urljoin

base = "https://example.com/docs/index.html"
print(urljoin(base, "guide.html"))
print(urljoin(base, "/api/v1"))
print(urljoin(base, "../images/logo.png"))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "https://example.com/docs/guide.html",
            "https://example.com/api/v1",
            "https://example.com/images/logo.png"
        ]
    );
}
