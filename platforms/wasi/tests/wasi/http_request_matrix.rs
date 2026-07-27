use std::io::{Read, Write};
use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use std::thread;

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-http-request-matrix-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn types(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http/types", name, args)
}

fn outgoing(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http/outgoing-handler", name, args)
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn is_error(value: &Value) -> Option<String> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(Value::String(text)) = object.properties.get("__wasi_error") {
            return Some(text.to_string());
        }
    }
    None
}

fn capture_request_line() -> (String, Arc<Mutex<Option<String>>>) {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind test server");
    let address = listener.local_addr().expect("server addr");
    let request_line = Arc::new(Mutex::new(None));
    let request_line_clone = request_line.clone();

    thread::spawn(move || {
        let (mut stream, _) = listener.accept().expect("accept request");
        let mut buffer = [0u8; 2048];
        let read = stream.read(&mut buffer).expect("read request");
        let raw = String::from_utf8_lossy(&buffer[..read]).to_string();
        let first_line = raw
            .lines()
            .next()
            .unwrap_or_default()
            .trim_end()
            .to_string();
        *request_line_clone.lock().unwrap() = Some(first_line);
        let response = "HTTP/1.1 200 OK\r\nContent-Length: 2\r\nContent-Type: text/plain\r\n\r\nok";
        stream
            .write_all(response.as_bytes())
            .expect("write response");
    });

    (format!("{}", address), request_line)
}

fn perform_request(method: Option<&str>, path: Option<&str>) -> String {
    let (authority, request_line) = capture_request_line();
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    if let Some(method) = method {
        assert!(
            is_error(&types(
                "[method]outgoing-request.set-method",
                vec![request.clone(), s(method)]
            ))
            .is_none()
        );
    }
    if let Some(path) = path {
        assert!(
            is_error(&types(
                "[method]outgoing-request.set-path-with-query",
                vec![request.clone(), s(path)]
            ))
            .is_none()
        );
    }
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    request_line.lock().unwrap().clone().unwrap_or_default()
}

macro_rules! method_case {
    ($name:ident, $method:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(perform_request(Some($method), Some("/matrix")), $expected);
        }
    };
}

macro_rules! path_case {
    ($name:ident, $path:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(perform_request(Some("GET"), Some($path)), $expected);
        }
    };
}

method_case!(
    request_line_uses_uppercase_get_method,
    "get",
    "GET /matrix HTTP/1.1"
);
method_case!(
    request_line_uses_uppercase_post_method,
    "post",
    "POST /matrix HTTP/1.1"
);
method_case!(
    request_line_uses_uppercase_put_method,
    "put",
    "PUT /matrix HTTP/1.1"
);
method_case!(
    request_line_uses_uppercase_patch_method,
    "patch",
    "PATCH /matrix HTTP/1.1"
);
method_case!(
    request_line_uses_uppercase_delete_method,
    "delete",
    "DELETE /matrix HTTP/1.1"
);
method_case!(
    request_line_uses_uppercase_head_method,
    "head",
    "HEAD /matrix HTTP/1.1"
);

path_case!(empty_path_string_normalizes_to_slash, "", "GET / HTTP/1.1");
path_case!(root_path_stays_root, "/", "GET / HTTP/1.1");
path_case!(
    relative_path_gains_leading_slash,
    "items",
    "GET /items HTTP/1.1"
);
path_case!(
    relative_path_with_query_gains_leading_slash,
    "items?page=1",
    "GET /items?page=1 HTTP/1.1"
);
path_case!(
    absolute_path_with_query_is_preserved,
    "/items?page=1",
    "GET /items?page=1 HTTP/1.1"
);

#[test]
fn default_request_line_is_get_slash_when_method_and_path_are_unset() {
    assert_eq!(perform_request(None, None), "GET / HTTP/1.1");
}
