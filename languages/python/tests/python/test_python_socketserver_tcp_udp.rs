use super::helpers::run_python;

// socketserver — TCPServer, UDPServer, BaseRequestHandler, StreamRequestHandler, DatagramRequestHandler, ThreadingMixIn

#[test]
fn test_socketserver_tcp_server_echo_roundtrip() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class EchoTCPHandler(socketserver.StreamRequestHandler):
    def handle(self):
        data = self.rfile.readline().strip()
        self.wfile.write(data + b"_ack\n")

server = socketserver.TCPServer(("127.0.0.1", 0), EchoTCPHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
client.connect((ip, port))
client.sendall(b"ping\n")
response = client.recv(1024).strip()
client.close()
server.server_close()
t.join()
print(response.decode())
"#,
    );
    assert_eq!(out, vec!["ping_ack"]);
}

#[test]
fn test_socketserver_udp_server_echo_roundtrip() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class EchoUDPHandler(socketserver.DatagramRequestHandler):
    def handle(self):
        data = self.rfile.read().strip()
        self.wfile.write(data + b"_udp_ack")

server = socketserver.UDPServer(("127.0.0.1", 0), EchoUDPHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
client.sendto(b"hello_udp", (ip, port))
response, _ = client.recvfrom(1024)
client.close()
server.server_close()
t.join()
print(response.decode())
"#,
    );
    assert_eq!(out, vec!["hello_udp_udp_ack"]);
}

#[test]
fn test_socketserver_threading_tcp_server() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class ThreadedTCPHandler(socketserver.BaseRequestHandler):
    def handle(self):
        data = self.request.recv(1024)
        cur_thread = threading.current_thread()
        self.request.sendall(f"thread_{cur_thread.name}".encode())

class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    pass

server = ThreadedTCPServer(("127.0.0.1", 0), ThreadedTCPHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
client.connect((ip, port))
client.sendall(b"req")
resp = client.recv(1024).decode()
client.close()
server.server_close()
t.join()
print(resp.startswith("thread_"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_socketserver_base_request_handler_attributes() {
    let out = run_python(
        r#"
import socketserver, socket, threading

captured = []

class InfoHandler(socketserver.BaseRequestHandler):
    def handle(self):
        captured.append(self.client_address[0])
        captured.append(isinstance(self.server, socketserver.TCPServer))
        self.request.sendall(b"ok")

server = socketserver.TCPServer(("127.0.0.1", 0), InfoHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.create_connection((ip, port))
c.recv(10)
c.close()
server.server_close()
t.join()
print(captured[0] in ("127.0.0.1", "localhost"))
print(captured[1])
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_socketserver_allow_reuse_address() {
    let out = run_python(
        r#"
import socketserver

class CustomServer(socketserver.TCPServer):
    allow_reuse_address = True

server = CustomServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
print(server.allow_reuse_address)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_socketserver_server_address_binding() {
    let out = run_python(
        r#"
import socketserver
server = socketserver.TCPServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
host, port = server.server_address
print(host)
print(port > 0)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["127.0.0.1", "True"]);
}

#[test]
fn test_socketserver_timeout_attribute() {
    let out = run_python(
        r#"
import socketserver
server = socketserver.TCPServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
server.timeout = 0.05
print(server.timeout)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["0.05"]);
}

#[test]
fn test_socketserver_handle_timeout_callback() {
    let out = run_python(
        r#"
import socketserver

class TimeoutServer(socketserver.TCPServer):
    timeout = 0.01
    timed_out = False
    def handle_timeout(self):
        self.timed_out = True

server = TimeoutServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
server.handle_request()
print(server.timed_out)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_socketserver_stream_request_handler_wfile_flush() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class FlushHandler(socketserver.StreamRequestHandler):
    def handle(self):
        self.wfile.write(b"flushed_data\n")
        self.wfile.flush()

server = socketserver.TCPServer(("127.0.0.1", 0), FlushHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.create_connection((ip, port))
data = c.recv(1024).strip()
c.close()
server.server_close()
t.join()
print(data.decode())
"#,
    );
    assert_eq!(out, vec!["flushed_data"]);
}

#[test]
fn test_socketserver_datagram_request_handler_socket_ref() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class RefHandler(socketserver.DatagramRequestHandler):
    def handle(self):
        self.wfile.write(b"ok")

server = socketserver.UDPServer(("127.0.0.1", 0), RefHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
c.sendto(b"query", (ip, port))
resp, _ = c.recvfrom(100)
c.close()
server.server_close()
t.join()
print(resp.decode())
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_socketserver_setup_and_finish_hooks() {
    let out = run_python(
        r#"
import socketserver, socket, threading

hooks = []

class HookHandler(socketserver.BaseRequestHandler):
    def setup(self):
        hooks.append("setup")
    def handle(self):
        hooks.append("handle")
    def finish(self):
        hooks.append("finish")

server = socketserver.TCPServer(("127.0.0.1", 0), HookHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.create_connection((ip, port))
c.close()
server.server_close()
t.join()
print(hooks)
"#,
    );
    assert_eq!(out, vec!["['setup', 'handle', 'finish']"]);
}

#[test]
fn test_socketserver_request_queue_size() {
    let out = run_python(
        r#"
import socketserver

class CustomServer(socketserver.TCPServer):
    request_queue_size = 10

server = CustomServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
print(server.request_queue_size)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_socketserver_fileno_method() {
    let out = run_python(
        r#"
import socketserver
server = socketserver.TCPServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
print(isinstance(server.fileno(), int))
server.server_close()
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_socketserver_verify_request() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class FilterServer(socketserver.TCPServer):
    def verify_request(self, request, client_address):
        return True

server = FilterServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.create_connection((ip, port))
c.close()
server.server_close()
t.join()
print("verified")
"#,
    );
    assert_eq!(out, vec!["verified"]);
}

#[test]
fn test_socketserver_shutdown_request() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class ShutdownHandler(socketserver.BaseRequestHandler):
    def handle(self):
        self.server.shutdown_request(self.request)

server = socketserver.TCPServer(("127.0.0.1", 0), ShutdownHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.create_connection((ip, port))
c.close()
server.server_close()
t.join()
print("shutdown_request_passed")
"#,
    );
    assert_eq!(out, vec!["shutdown_request_passed"]);
}

#[test]
fn test_socketserver_threading_mixin_daemon_threads() {
    let out = run_python(
        r#"
import socketserver

class DaemonTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    daemon_threads = True

server = DaemonTCPServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
print(server.daemon_threads)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_socketserver_close_request() {
    let out = run_python(
        r#"
import socketserver, socket

server = socketserver.TCPServer(("127.0.0.1", 0), socketserver.BaseRequestHandler)
ip, port = server.server_address
c = socket.create_connection((ip, port))
server.close_request(c)
server.server_close()
print("closed")
"#,
    );
    assert_eq!(out, vec!["closed"]);
}

#[test]
fn test_socketserver_udp_server_bind() {
    let out = run_python(
        r#"
import socketserver
server = socketserver.UDPServer(("127.0.0.1", 0), socketserver.DatagramRequestHandler)
ip, port = server.server_address
print(ip == "127.0.0.1")
print(port > 0)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_socketserver_max_packet_size_udp() {
    let out = run_python(
        r#"
import socketserver
class UDP(socketserver.UDPServer):
    max_packet_size = 4096

server = UDP(("127.0.0.1", 0), socketserver.DatagramRequestHandler)
print(server.max_packet_size)
server.server_close()
"#,
    );
    assert_eq!(out, vec!["4096"]);
}

#[test]
fn test_socketserver_handle_error_override() {
    let out = run_python(
        r#"
import socketserver, socket, threading

class ErrorHandlerServer(socketserver.TCPServer):
    error_caught = False
    def handle_error(self, request, client_address):
        self.error_caught = True

class FaultyHandler(socketserver.BaseRequestHandler):
    def handle(self):
        raise ValueError("simulated handler crash")

server = ErrorHandlerServer(("127.0.0.1", 0), FaultyHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

c = socket.create_connection((ip, port))
c.close()
server.server_close()
t.join()
print(server.error_caught)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
