use std::io::{Read, Write};
use std::net::TcpListener;
use std::thread;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-http-status-matrix-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn types(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http/types", name, args)
}

/// 0.3.1 replaced `outgoing-handler.handle` with `client.send`, which answers
/// the response directly instead of a `future-incoming-response`.
fn client(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http/client", name, args)
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
    let request = types("[static]request.new", vec![headers]);
    assert!(
        is_error(&types(
            "[method]request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    let response = client("send", vec![request]);
    types("[method]response.get-status-code", vec![response]).as_f64()
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
