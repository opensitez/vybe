use vybe_bytecode::value::{ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-cli-test>");
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
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn invoke_result(module: &str, name: &str, args: Vec<Value>) -> Result<Value, String> {
    let mut chunk = Chunk::new("<wasi-cli-test>");
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
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).map_err(|error| error.message)
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn array_len(value: &Value) -> usize {
    let Value::Object(object) = value else {
        return 0;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        return 0;
    };
    values.len()
}

fn array_strings(value: &Value) -> Vec<String> {
    let Value::Object(object) = value else {
        return Vec::new();
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        return Vec::new();
    };
    values
        .iter()
        .filter_map(|value| match value {
            Value::String(text) => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

// ── wasi:cli/environment ─────────────────────────────────────────────────────

#[test]
fn get_environment_returns_array_of_string_pairs() {
    let result = invoke("wasi:cli/environment", "get-environment", vec![]);
    if let Value::Object(object) = &result {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(entries) = &object.kind {
            for entry in entries {
                let Value::Object(pair) = entry else { continue };
                let pair = pair.lock().unwrap();
                let ObjectKind::Array(values) = &pair.kind else {
                    continue;
                };
                assert_eq!(values.len(), 2);
                assert!(matches!(values.first(), Some(Value::String(_))));
                assert!(matches!(values.get(1), Some(Value::String(_))));
                return;
            }
            return;
        }
    }
    panic!(
        "get-environment should return list<(string, string)>, got {:?}",
        result
    );
}

#[test]
fn get_environment_with_key_returns_value() {
    let key = if std::env::var("HOME").is_ok() {
        "HOME"
    } else {
        "USERPROFILE"
    };
    let expected = std::env::var(key).unwrap_or_else(|_| String::from("."));
    let result = invoke("wasi:cli/environment", "get-environment", vec![s(key)]);
    assert_eq!(result, s(&expected));
}

#[test]
fn get_environment_with_missing_key_returns_null() {
    let result = invoke(
        "wasi:cli/environment",
        "get-environment",
        vec![s("VYBE_TEST_ENV_SHOULD_NOT_EXIST")],
    );
    assert!(matches!(result, Value::Null));
}

#[test]
fn get_arguments_returns_string_array() {
    let result = invoke("wasi:cli/environment", "get-arguments", vec![]);
    assert!(array_len(&result) > 0);
}

#[test]
fn get_arguments_matches_process_argument_vector_exactly() {
    let result = invoke("wasi:cli/environment", "get-arguments", vec![]);
    let expected: Vec<String> = std::env::args().collect();
    assert_eq!(array_strings(&result), expected);
}

#[test]
fn initial_cwd_matches_get_initial_cwd() {
    let initial = invoke("wasi:cli/environment", "initial-cwd", vec![]);
    let renamed = invoke("wasi:cli/environment", "get-initial-cwd", vec![]);
    assert_eq!(initial, renamed);
    assert_eq!(
        initial,
        s(std::env::current_dir().unwrap().to_string_lossy().as_ref())
    );
}

#[test]
fn get_initial_cwd_returns_non_empty_string() {
    let result = invoke("wasi:cli/environment", "get-initial-cwd", vec![]);
    let Value::String(path) = result else {
        panic!("get-initial-cwd should return string for current process cwd");
    };
    assert!(!path.is_empty());
}

#[test]
fn get_environment_contains_home_or_userprofile_when_present() {
    let key = if std::env::var("HOME").is_ok() {
        "HOME"
    } else {
        "USERPROFILE"
    };
    let expected = std::env::var(key).unwrap_or_else(|_| String::from("."));
    let result = invoke("wasi:cli/environment", "get-environment", vec![]);

    let Value::Object(object) = result else {
        panic!("get-environment should return array");
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(entries) = &object.kind else {
        panic!("get-environment should return array");
    };

    let found = entries.iter().any(|entry| {
        let Value::Object(pair) = entry else {
            return false;
        };
        let pair = pair.lock().unwrap();
        let ObjectKind::Array(values) = &pair.kind else {
            return false;
        };
        matches!(values.first(), Some(Value::String(name)) if name.as_ref() == key)
            && matches!(values.get(1), Some(Value::String(value)) if value.as_ref() == expected)
    });

    assert!(found, "expected get-environment to include {key}");
}

// ── wasi:logging/logging ─────────────────────────────────────────────────────

#[test]
fn logging_log_1_arg_returns_null() {
    let result = invoke("wasi:logging/logging", "log", vec![s("hello world")]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn logging_log_2_args_level_and_message_returns_null() {
    let result = invoke(
        "wasi:logging/logging",
        "log",
        vec![s("error"), s("something failed")],
    );
    assert!(matches!(result, Value::Null));
}

#[test]
fn logging_log_3_args_level_context_message_returns_null() {
    let result = invoke(
        "wasi:logging/logging",
        "log",
        vec![s("warn"), s("mymodule"), s("low disk")],
    );
    assert!(matches!(result, Value::Null));
}

#[test]
fn logging_log_0_args_returns_null() {
    let result = invoke("wasi:logging/logging", "log", vec![]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn logging_log_variadic_returns_null() {
    let result = invoke(
        "wasi:logging/logging",
        "log",
        vec![s("a"), s("b"), s("c"), s("d")],
    );
    assert!(matches!(result, Value::Null));
}

// ── wasi:cli/exit ────────────────────────────────────────────────────────────

#[test]
fn proposal_cli_exit_import_is_registered() {
    assert!(
        has_import("wasi:cli/exit", "exit"),
        "wasi:cli/exit.exit should be covered by the CLI category"
    );
}

// ── wasi:cli/stdout|stderr|stdin ─────────────────────────────────────────────

#[test]
fn proposal_cli_stdout_get_stdout_returns_stream_handle() {
    let result = invoke("wasi:cli/stdout", "get-stdout", vec![]);
    let Value::Object(obj) = result else {
        panic!("get-stdout should return a stream handle object");
    };
    let obj = obj.lock().unwrap();
    assert_eq!(obj.properties.get("fd"), Some(&Value::I32(1)));
}

#[test]
fn proposal_cli_stderr_get_stderr_returns_stream_handle() {
    let result = invoke("wasi:cli/stderr", "get-stderr", vec![]);
    let Value::Object(obj) = result else {
        panic!("get-stderr should return a stream handle object");
    };
    let obj = obj.lock().unwrap();
    assert_eq!(obj.properties.get("fd"), Some(&Value::I32(2)));
}

#[test]
fn proposal_cli_stdin_get_stdin_returns_stream_handle() {
    let result = invoke("wasi:cli/stdin", "get-stdin", vec![]);
    let Value::Object(obj) = result else {
        panic!("get-stdin should return a stream handle object");
    };
    let obj = obj.lock().unwrap();
    assert_eq!(obj.properties.get("fd"), Some(&Value::I32(0)));
}

#[test]
fn proposal_cli_stdout_write_via_stream_import_resolves() {
    assert!(
        invoke_result("wasi:cli/stdout", "write-via-stream", vec![Value::Null]).is_ok(),
        "wasi:cli/stdout.write-via-stream should be covered by the CLI category"
    );
}

#[test]
fn proposal_cli_stderr_write_via_stream_import_resolves() {
    assert!(
        invoke_result("wasi:cli/stderr", "write-via-stream", vec![Value::Null]).is_ok(),
        "wasi:cli/stderr.write-via-stream should be covered by the CLI category"
    );
}
