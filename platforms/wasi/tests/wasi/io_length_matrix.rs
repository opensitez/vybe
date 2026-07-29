use std::net::TcpListener;
use std::sync::{Arc, Mutex};

use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-io-length-matrix-test>");
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

fn len_of(value: &Value) -> usize {
    let Value::Object(object) = value else {
        return 0;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return 0;
    };
    elements.len()
}

fn element_at(value: &Value, index: usize) -> Value {
    let Value::Object(object) = value else {
        return Value::Null;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return Value::Null;
    };
    elements.get(index).cloned().unwrap_or(Value::Null)
}

fn bytes(values: &[u8]) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(
        values
            .iter()
            .map(|value| Value::I32(*value as i32))
            .collect(),
    ))))
}

fn bytes_to_vec(value: &Value) -> Vec<u8> {
    let Value::Object(object) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return Vec::new();
    };
    elements
        .iter()
        .map(|value| value.as_i32().clamp(0, 255) as u8)
        .collect()
}

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

fn connected_tcp_streams() -> (Value, Value, Value, Value) {
    let port = free_port();
    let addr = Value::String(Arc::from(format!("127.0.0.1:{}", port).as_str()));
    let family = Value::String(Arc::from("ipv4"));

    let network = call_import("wasi:sockets/instance-network", "instance-network", vec![]);
    let listener = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family.clone()],
    );
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-bind",
            vec![listener.clone(), network.clone(), addr.clone()]
        )
        .as_bool(),
        true
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-bind", vec![listener.clone()]).as_bool(),
        true
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "start-listen", vec![listener.clone()]).as_bool(),
        true
    );
    assert_eq!(
        call_import("wasi:sockets/tcp", "finish-listen", vec![listener.clone()]).as_bool(),
        true
    );

    let client = call_import(
        "wasi:sockets/tcp-create-socket",
        "create-tcp-socket",
        vec![family],
    );
    assert_eq!(
        call_import(
            "wasi:sockets/tcp",
            "start-connect",
            vec![client.clone(), network, addr]
        )
        .as_bool(),
        true
    );

    let client_streams = call_import("wasi:sockets/tcp", "finish-connect", vec![client]);
    let client_in = element_at(&client_streams, 0);
    let client_out = element_at(&client_streams, 1);

    let accepted = {
        let mut last = Value::Null;
        for _ in 0..100 {
            last = call_import("wasi:sockets/tcp", "accept", vec![listener.clone()]);
            if len_of(&last) == 3 {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        last
    };

    let server_in = element_at(&accepted, 1);
    let server_out = element_at(&accepted, 2);
    (client_in, client_out, server_in, server_out)
}

macro_rules! read_length_case {
    ($name:ident, $length:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let (_, client_out, server_in, _) = connected_tcp_streams();
            assert!(matches!(
                call_import("wasi:io/streams", "write", vec![client_out, bytes(b"abcd")]),
                Value::Null
            ));
            let result = call_import(
                "wasi:io/streams",
                "read",
                vec![server_in, Value::I64($length)],
            );
            assert_eq!(bytes_to_vec(&result), $expected);
        }
    };
}

macro_rules! blocking_read_length_case {
    ($name:ident, $length:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let (_, client_out, server_in, _) = connected_tcp_streams();
            assert!(matches!(
                call_import("wasi:io/streams", "write", vec![client_out, bytes(b"abcd")]),
                Value::Null
            ));
            let result = call_import(
                "wasi:io/streams",
                "blocking-read",
                vec![server_in, Value::I64($length)],
            );
            assert_eq!(bytes_to_vec(&result), $expected);
        }
    };
}

macro_rules! write_zeroes_length_case {
    ($name:ident, $method:expr, $length:expr) => {
        #[test]
        fn $name() {
            let (client_in, _, _, server_out) = connected_tcp_streams();
            assert!(matches!(
                call_import(
                    "wasi:io/streams",
                    $method,
                    vec![server_out, Value::I64($length)]
                ),
                Value::Null
            ));
            let result = call_import(
                "wasi:io/streams",
                "blocking-read",
                vec![client_in, Value::I64($length)],
            );
            assert_eq!(bytes_to_vec(&result), vec![0; $length as usize]);
        }
    };
}

macro_rules! splice_length_case {
    ($name:ident, $method:expr, $length:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let (_, source_out, source_in, _) = connected_tcp_streams();
            let (target_in, _, _, target_out) = connected_tcp_streams();
            assert!(matches!(
                call_import("wasi:io/streams", "write", vec![source_out, bytes(b"abcd")]),
                Value::Null
            ));
            let moved = call_import(
                "wasi:io/streams",
                $method,
                vec![target_out, source_in, Value::I64($length)],
            );
            assert_eq!(moved.as_i64(), $expected.len() as i64);
            let copied = call_import(
                "wasi:io/streams",
                "blocking-read",
                vec![target_in, Value::I64($expected.len() as i64)],
            );
            assert_eq!(bytes_to_vec(&copied), $expected);
        }
    };
}

read_length_case!(read_length_one_returns_single_byte, 1, b"a");
read_length_case!(read_length_four_returns_full_payload, 4, b"abcd");
read_length_case!(read_length_eight_returns_available_payload, 8, b"abcd");

blocking_read_length_case!(blocking_read_length_one_returns_single_byte, 1, b"a");
blocking_read_length_case!(blocking_read_length_four_returns_full_payload, 4, b"abcd");
blocking_read_length_case!(
    blocking_read_length_eight_returns_available_payload,
    8,
    b"abcd"
);

write_zeroes_length_case!(write_zeroes_length_one_emits_single_zero, "write-zeroes", 1);
write_zeroes_length_case!(
    write_zeroes_length_four_emits_four_zeroes,
    "write-zeroes",
    4
);
write_zeroes_length_case!(
    write_zeroes_length_eight_emits_eight_zeroes,
    "write-zeroes",
    8
);

write_zeroes_length_case!(
    blocking_write_zeroes_length_one_emits_single_zero,
    "blocking-write-zeroes-and-flush",
    1
);
write_zeroes_length_case!(
    blocking_write_zeroes_length_four_emits_four_zeroes,
    "blocking-write-zeroes-and-flush",
    4
);
write_zeroes_length_case!(
    blocking_write_zeroes_length_eight_emits_eight_zeroes,
    "blocking-write-zeroes-and-flush",
    8
);

splice_length_case!(splice_length_one_moves_single_byte, "splice", 1, b"a");
splice_length_case!(splice_length_four_moves_full_payload, "splice", 4, b"abcd");
splice_length_case!(
    splice_length_eight_moves_available_payload,
    "splice",
    8,
    b"abcd"
);

splice_length_case!(
    blocking_splice_length_one_moves_single_byte,
    "blocking-splice",
    1,
    b"a"
);
splice_length_case!(
    blocking_splice_length_four_moves_full_payload,
    "blocking-splice",
    4,
    b"abcd"
);
splice_length_case!(
    blocking_splice_length_eight_moves_available_payload,
    "blocking-splice",
    8,
    b"abcd"
);
