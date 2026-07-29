use std::io::{Read, Write};
use std::net::TcpListener;
use std::thread;

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-http-status-matrix-test>");
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

fn legacy(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http", name, args)
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

fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(value) = object.properties.get(key) {
            return value.clone();
        }
    }
    Value::Null
}

fn start_server(status_line: &str, body: &str) -> String {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind test server");
    let address = listener.local_addr().expect("server addr");
    let status_line = status_line.to_string();
    let body = body.to_string();

    thread::spawn(move || {
        let (mut stream, _) = listener.accept().expect("accept request");
        let mut buffer = [0u8; 1024];
        let _ = stream.read(&mut buffer);
        let response = format!(
            "HTTP/1.1 {}\r\nContent-Length: {}\r\nContent-Type: text/plain\r\n\r\n{}",
            status_line,
            body.len(),
            body,
        );
        stream
            .write_all(response.as_bytes())
            .expect("write response");
    });

    format!("{}", address)
}

fn request_status(status_line: &str, body: &str) -> f64 {
    let authority = start_server(status_line, body);
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    types("[method]incoming-response.status", vec![response]).as_f64()
}

macro_rules! outgoing_status_test {
    ($name:ident, $status_line:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(request_status($status_line, "body"), $expected);
        }
    };
}

outgoing_status_test!(outgoing_status_reports_ok_200, "200 OK", 200.0);
outgoing_status_test!(outgoing_status_reports_created_201, "201 Created", 201.0);
outgoing_status_test!(
    outgoing_status_reports_no_content_204,
    "204 No Content",
    204.0
);
outgoing_status_test!(
    outgoing_status_reports_moved_permanently_301,
    "301 Moved Permanently",
    301.0
);
outgoing_status_test!(outgoing_status_reports_found_302, "302 Found", 302.0);
outgoing_status_test!(
    outgoing_status_reports_not_found_404,
    "404 Not Found",
    404.0
);
outgoing_status_test!(
    outgoing_status_reports_internal_server_error_500,
    "500 Internal Server Error",
    500.0
);
outgoing_status_test!(
    outgoing_status_reports_service_unavailable_503,
    "503 Service Unavailable",
    503.0
);
