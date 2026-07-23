use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Socket Networking & Protocols — socket, AF_INET, SOCK_STREAM, SOCK_DGRAM, getsockname, setsockopt, SO_REUSEADDR, timeout
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_socket_hostname_and_address_resolution() {
    let src = r#"
import socket

hostname = socket.gethostname()
print(isinstance(hostname, str) and len(hostname) > 0)
ip = socket.gethostbyname("localhost")
print(ip in ("127.0.0.1", "::1"))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_socket_inet_aton_and_ntoa_conversions() {
    let src = r#"
import socket

ip_str = "127.0.0.1"
packed = socket.inet_aton(ip_str)
print(len(packed))  # 4 bytes
print(socket.inet_ntoa(packed) == ip_str)
"#;
    assert_eq!(run_python(src), vec!["4", "True"]);
}

#[test]
fn test_py_socket_tcp_loopback_server_client() {
    let src = r#"
import socket, threading

server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server.bind(("127.0.0.1", 0))
port = server.getsockname()[1]
server.listen(1)

def handle_client():
    conn, _ = server.accept()
    data = conn.recv(1024)
    conn.sendall(b"REPLY: " + data)
    conn.close()

t = threading.Thread(target=handle_client)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
client.connect(("127.0.0.1", port))
client.sendall(b"PING")
res = client.recv(1024)
client.close()
t.join()
server.close()

print(res.decode())
"#;
    assert_eq!(run_python(src), vec!["REPLY: PING"]);
}

#[test]
fn test_py_socket_udp_loopback_communication() {
    let src = r#"
import socket, threading

server = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
server.bind(("127.0.0.1", 0))
port = server.getsockname()[1]

def udp_echo():
    data, addr = server.recvfrom(1024)
    server.sendto(b"UDP_ACK: " + data, addr)

t = threading.Thread(target=udp_echo)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
client.sendto(b"PACKET", ("127.0.0.1", port))
reply, _ = client.recvfrom(1024)
client.close()
t.join()
server.close()

print(reply.decode())
"#;
    assert_eq!(run_python(src), vec!["UDP_ACK: PACKET"]);
}

#[test]
fn test_py_socket_timeout_configuration() {
    let src = r#"
import socket

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.settimeout(0.05)
print(s.gettimeout())
s.close()
"#;
    assert_eq!(run_python(src), vec!["0.05"]);
}

#[test]
fn test_py_socket_setsockopt_reuseaddr() {
    let src = r#"
import socket

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
val = s.getsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR)
print(val != 0)
s.close()
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_socket_create_connection_helper() {
    let src = r#"
import socket, threading

server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server.bind(("127.0.0.1", 0))
port = server.getsockname()[1]
server.listen(1)

def handle():
    conn, _ = server.accept()
    conn.close()

t = threading.Thread(target=handle)
t.start()

client = socket.create_connection(("127.0.0.1", port))
print(client.fileno() > 0)
client.close()
t.join()
server.close()
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_socket_shutdown_modes() {
    let src = r#"
import socket

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
print(hasattr(socket, "SHUT_RD"))
print(hasattr(socket, "SHUT_WR"))
print(hasattr(socket, "SHUT_RDWR"))
s.close()
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_socket_getservbyname_well_known_ports() {
    let src = r#"
import socket

print(socket.getservbyname("http"))
print(socket.getservbyname("https"))
"#;
    assert_eq!(run_python(src), vec!["80", "443"]);
}

#[test]
fn test_py_socket_getaddrinfo_resolution() {
    let src = r#"
import socket

info = socket.getaddrinfo("localhost", 80, family=socket.AF_INET, type=socket.SOCK_STREAM)
print(len(info) > 0)
family, socktype, proto, canonname, sockaddr = info[0]
print(family == socket.AF_INET)
print(sockaddr[0] == "127.0.0.1")
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}
