use super::helpers::run_python;

// http.client — HTTPConnection, HTTPResponse, HTTPMessage, responses, status codes, request methods, header parsing

#[test]
fn test_http_client_responses_dict_status_lookup() {
    let out = run_python(r#"
import http.client
print(http.client.responses[200])
print(http.client.responses[404])
print(http.client.responses[500])
"#);
    assert_eq!(out, vec!["OK", "Not Found", "Internal Server Error"]);
}

#[test]
fn test_http_client_status_code_constants() {
    let out = run_python(r#"
import http.client
print(http.client.OK)
print(http.client.NOT_FOUND)
print(http.client.FOUND)
"#);
    assert_eq!(out, vec!["200", "404", "302"]);
}

#[test]
fn test_http_client_parse_headers_from_file_like() {
    let out = run_python(r#"
import http.client, io
header_bytes = b"Content-Type: text/html\r\nContent-Length: 100\r\n\r\n"
buf = io.BytesIO(header_bytes)
msg = http.client.parse_headers(buf)
print(msg.get_content_type())
print(msg["content-length"])
"#);
    assert_eq!(out, vec!["text/html", "100"]);
}

#[test]
fn test_http_client_connection_initialization() {
    let out = run_python(r#"
import http.client
conn = http.client.HTTPConnection("example.com", 80, timeout=10)
print(conn.host)
print(conn.port)
print(conn.timeout)
conn.close()
"#);
    assert_eq!(out, vec!["example.com", "80", "10"]);
}

#[test]
fn test_http_client_http_exception_hierarchy() {
    let out = run_python(r#"
import http.client
print(issubclass(http.client.BadStatusLine, http.client.HTMLElement))
print(issubclass(http.client.CannotSendRequest, http.client.HTTPException))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_http_client_tunnel_set_proxy() {
    let out = run_python(r#"
import http.client
conn = http.client.HTTPConnection("proxy.local", 8080)
conn.set_tunnel("target.local", 443, headers={"Proxy-Authorization": "Basic xyz"})
print(conn._tunnel_host)
print(conn._tunnel_port)
conn.close()
"#);
    assert_eq!(out, vec!["target.local", "443"]);
}

#[test]
fn test_http_client_response_read_chunks() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp_data = b"HTTP/1.1 200 OK\r\nContent-Length: 11\r\n\r\nHello World"
sock = DummySocket(resp_data)
resp = http.client.HTTPResponse(sock)
resp.begin()
print(resp.status)
print(resp.reason)
print(resp.read().decode())
"#);
    assert_eq!(out, vec!["200", "OK", "Hello World"]);
}

#[test]
fn test_http_client_response_getheader_case_insensitive() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp_data = b"HTTP/1.1 200 OK\r\nX-Custom-Header: TestVal\r\nContent-Length: 0\r\n\r\n"
resp = http.client.HTTPResponse(DummySocket(resp_data))
resp.begin()
print(resp.getheader("x-custom-header"))
print(resp.getheader("X-CUSTOM-HEADER"))
"#);
    assert_eq!(out, vec!["TestVal", "TestVal"]);
}

#[test]
fn test_http_client_response_getheaders_list() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp_data = b"HTTP/1.1 200 OK\r\nServer: Dummy\r\nContent-Length: 0\r\n\r\n"
resp = http.client.HTTPResponse(DummySocket(resp_data))
resp.begin()
headers = dict(resp.getheaders())
print(headers["Server"])
"#);
    assert_eq!(out, vec!["Dummy"]);
}

#[test]
fn test_http_client_chunked_transfer_encoding_reading() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp_data = b"HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n5\r\nhello\r\n6\r\n world\r\n0\r\n\r\n"
resp = http.client.HTTPResponse(DummySocket(resp_data))
resp.begin()
print(resp.read().decode())
"#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn test_http_client_response_isclosed() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp = http.client.HTTPResponse(DummySocket(b"HTTP/1.1 200 OK\r\nContent-Length: 0\r\n\r\n"))
resp.begin()
print(resp.isclosed())
resp.close()
print(resp.isclosed())
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_http_client_connection_auto_open() {
    let out = run_python(r#"
import http.client
conn = http.client.HTTPConnection("localhost", 9999)
print(conn.sock is None)
conn.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_http_client_putrequest_and_putheader() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self): self.buf = io.BytesIO()
    def sendall(self, data): self.buf.write(data)
    def close(self): pass

conn = http.client.HTTPConnection("dummy.host", 80)
conn.sock = DummySocket()
conn.putrequest("GET", "/api/v1/users")
conn.putheader("Accept", "application/json")
conn.endheaders()
data = conn.sock.buf.getvalue().decode()
print("GET /api/v1/users HTTP/1.1" in data)
print("Accept: application/json" in data)
conn.close()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_http_client_request_with_body_bytes() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self): self.buf = io.BytesIO()
    def sendall(self, data): self.buf.write(data)
    def close(self): pass

conn = http.client.HTTPConnection("dummy.host", 80)
conn.sock = DummySocket()
conn.request("POST", "/submit", body=b"payload_data", headers={"Content-Type": "text/plain"})
sent = conn.sock.buf.getvalue().decode()
print("Content-Length: 12" in sent)
print("payload_data" in sent)
conn.close()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_http_client_response_version_attribute() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp = http.client.HTTPResponse(DummySocket(b"HTTP/1.1 200 OK\r\nContent-Length: 0\r\n\r\n"))
resp.begin()
print(resp.version)  # 11 for HTTP/1.1
"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_http_client_cannot_send_request_state() {
    let out = run_python(r#"
import http.client
conn = http.client.HTTPConnection("dummy.host", 80)
conn._state = http.client._CS_REQ_SENT
try:
    conn.putrequest("GET", "/")
except http.client.CannotSendRequest:
    print("CannotSendRequest")
conn.close()
"#);
    assert_eq!(out, vec!["CannotSendRequest"]);
}

#[test]
fn test_http_client_response_readline_streaming() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

body = b"HTTP/1.1 200 OK\r\n\r\nline1\nline2\nline3\n"
resp = http.client.HTTPResponse(DummySocket(body))
resp.begin()
print(resp.readline().decode().strip())
print(resp.readline().decode().strip())
"#);
    assert_eq!(out, vec!["line1", "line2"]);
}

#[test]
fn test_http_client_response_readinto_buffer() {
    let out = run_python(r#"
import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp = http.client.HTTPResponse(DummySocket(b"HTTP/1.1 200 OK\r\nContent-Length: 4\r\n\r\nABCD"))
resp.begin()
buf = bytearray(4)
n = resp.readinto(buf)
print(n)
print(buf.decode())
"#);
    assert_eq!(out, vec!["4", "ABCD"]);
}

#[test]
fn test_http_client_unimplemented_scheme_raises() {
    let out = run_python(r#"
import http.client
try:
    http.client.HTTPConnection("http://invalid_format:80")
except Exception:
    print("Error")
"#);
    assert_eq!(out, vec!["Error"]);
}

#[test]
fn test_http_client_status_codes_informational() {
    let out = run_python(r#"
import http.client
print(http.client.CONTINUE)
print(http.client.SWITCHING_PROTOCOLS)
"#);
    assert_eq!(out, vec!["100", "101"]);
}
