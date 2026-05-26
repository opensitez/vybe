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
    vm.host_registry.contains_key(&(module.to_string(), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn array_len(value: &Value) -> usize {
    let Value::Object(object) = value else { return 0 };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { return 0 };
    values.len()
}

fn array_strings(value: &Value) -> Vec<String> {
    let Value::Object(object) = value else { return Vec::new() };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { return Vec::new() };
    values
        .iter()
        .filter_map(|value| match value {
            Value::String(text) => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

#[test]
fn get_environment_returns_array_of_string_pairs() {
    let result = invoke("wasi:cli/environment", "get-environment", vec![]);
    if let Value::Object(object) = &result {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(entries) = &object.kind {
            for entry in entries {
                let Value::Object(pair) = entry else { continue };
                let pair = pair.lock().unwrap();
                let ObjectKind::Array(values) = &pair.kind else { continue };
                assert_eq!(values.len(), 2);
                assert!(matches!(values.first(), Some(Value::String(_))));
                assert!(matches!(values.get(1), Some(Value::String(_))));
                return;
            }
            return;
        }
    }
    panic!("get-environment should return list<(string, string)>, got {:?}", result);
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
fn get_arguments_matches_legacy_args_surface() {
    let proposal = invoke("wasi:cli/environment", "get-arguments", vec![]);
    let legacy = invoke("wasi:cli", "args", vec![]);
    assert_eq!(array_strings(&proposal), array_strings(&legacy));
}

#[test]
fn initial_cwd_matches_get_initial_cwd() {
    let initial = invoke("wasi:cli/environment", "initial-cwd", vec![]);
    let renamed = invoke("wasi:cli/environment", "get-initial-cwd", vec![]);
    assert_eq!(initial, renamed);
    assert_eq!(initial, s(std::env::current_dir().unwrap().to_string_lossy().as_ref()));
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
    let key = if std::env::var("HOME").is_ok() { "HOME" } else { "USERPROFILE" };
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
        let Value::Object(pair) = entry else { return false };
        let pair = pair.lock().unwrap();
        let ObjectKind::Array(values) = &pair.kind else { return false };
        matches!(values.first(), Some(Value::String(name)) if name.as_ref() == key)
            && matches!(values.get(1), Some(Value::String(value)) if value.as_ref() == expected)
    });

    assert!(found, "expected get-environment to include {key}");
}

#[test]
fn cli_args_returns_string_array() {
    let result = invoke("wasi:cli", "args", vec![]);
    assert!(array_len(&result) > 0);
}

#[test]
fn cli_args_match_process_argument_vector_exactly() {
    let result = invoke("wasi:cli", "args", vec![]);
    let expected: Vec<String> = std::env::args().collect();
    assert_eq!(array_strings(&result), expected);
}

#[test]
fn cli_get_env_returns_null_for_missing_key() {
    let result = invoke("wasi:cli", "getEnv", vec![s("VYBE_TEST_ENV_SHOULD_NOT_EXIST")]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn cli_get_env_returns_home_when_present() {
    let key = if std::env::var("HOME").is_ok() { "HOME" } else { "USERPROFILE" };
    let expected = std::env::var(key).unwrap_or_else(|_| String::from("."));
    let result = invoke("wasi:cli", "getEnv", vec![s(key)]);
    assert_eq!(result, s(&expected));
}

#[test]
fn cli_cwd_matches_current_dir() {
    let result = invoke("wasi:cli", "cwd", vec![]);
    assert_eq!(result, s(std::env::current_dir().unwrap().to_string_lossy().as_ref()));
}

#[test]
fn cli_cwd_matches_proposal_initial_cwd_surfaces() {
    let cwd = invoke("wasi:cli", "cwd", vec![]);
    let initial = invoke("wasi:cli/environment", "initial-cwd", vec![]);
    let renamed = invoke("wasi:cli/environment", "get-initial-cwd", vec![]);
    assert_eq!(cwd, initial);
    assert_eq!(cwd, renamed);
}

#[test]
fn cli_platform_matches_std_env() {
    let result = invoke("wasi:cli", "platform", vec![]);
    assert_eq!(result, s(std::env::consts::OS));
}

#[test]
fn cli_arch_matches_std_env() {
    let result = invoke("wasi:cli", "arch", vec![]);
    assert_eq!(result, s(std::env::consts::ARCH));
}

#[test]
fn cli_machine_name_returns_string() {
    let result = invoke("wasi:cli", "machineName", vec![]);
    assert!(matches!(result, Value::String(_)));
}

#[test]
fn cli_machine_name_matches_environment_fallback_chain() {
    let expected = std::env::var("HOSTNAME")
        .or_else(|_| std::env::var("COMPUTERNAME"))
        .unwrap_or_else(|_| String::from("unknown"));
    let result = invoke("wasi:cli", "machineName", vec![]);
    assert_eq!(result, s(&expected));
}

#[test]
fn cli_user_name_returns_string() {
    let result = invoke("wasi:cli", "userName", vec![]);
    assert!(matches!(result, Value::String(_)));
}

#[test]
fn cli_user_name_matches_environment_fallback_chain() {
    let expected = std::env::var("USER")
        .or_else(|_| std::env::var("USERNAME"))
        .unwrap_or_else(|_| String::from("unknown"));
    let result = invoke("wasi:cli", "userName", vec![]);
    assert_eq!(result, s(&expected));
}

#[test]
fn cli_new_line_returns_line_feed() {
    let result = invoke("wasi:cli", "newLine", vec![]);
    assert_eq!(result, s("\n"));
}

#[test]
fn cli_tick_count_is_non_negative() {
    let result = invoke("wasi:cli", "tickCount", vec![]);
    let Value::F64(number) = result else {
        panic!("tickCount should return f64");
    };
    assert!(number >= 0.0);
}

#[test]
fn cli_tick_count_is_non_decreasing() {
    let first = invoke("wasi:cli", "tickCount", vec![]).as_f64();
    let second = invoke("wasi:cli", "tickCount", vec![]).as_f64();
    assert!(second >= first);
}

#[test]
fn cli_get_folder_path_returns_string() {
    let result = invoke("wasi:cli", "getFolderPath", vec![]);
    assert!(matches!(result, Value::String(_)));
}

#[test]
fn cli_get_folder_path_uses_home_directory_fallback() {
    let expected = std::env::var("HOME")
        .or_else(|_| std::env::var("USERPROFILE"))
        .unwrap_or_else(|_| String::from("."));
    let result = invoke("wasi:cli", "getFolderPath", vec![]);
    assert_eq!(result, s(&expected));
}

#[test]
fn cli_log_returns_null() {
    let result = invoke("wasi:cli", "log", vec![s("hello"), s("world")]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn cli_warn_returns_null() {
    let result = invoke("wasi:cli", "warn", vec![s("warning")]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn cli_error_returns_null() {
    let result = invoke("wasi:cli", "error", vec![s("error")]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn cli_time_and_time_end_return_null() {
    let start = invoke("wasi:cli", "time", vec![s("wasi-cli-test")]);
    let end = invoke("wasi:cli", "timeEnd", vec![s("wasi-cli-test")]);
    assert!(matches!(start, Value::Null));
    assert!(matches!(end, Value::Null));
}

#[test]
fn cli_time_and_time_end_support_default_label() {
    let start = invoke("wasi:cli", "time", vec![]);
    let end = invoke("wasi:cli", "timeEnd", vec![]);
    assert!(matches!(start, Value::Null));
    assert!(matches!(end, Value::Null));
}

#[test]
fn cli_time_end_without_prior_time_is_still_null() {
    let result = invoke("wasi:cli", "timeEnd", vec![s("never-started")]);
    assert!(matches!(result, Value::Null));
}

#[test]
fn cli_log_warn_and_error_accept_empty_argument_lists() {
    assert!(matches!(invoke("wasi:cli", "log", vec![]), Value::Null));
    assert!(matches!(invoke("wasi:cli", "warn", vec![]), Value::Null));
    assert!(matches!(invoke("wasi:cli", "error", vec![]), Value::Null));
}

#[test]
fn proposal_cli_stdin_read_via_stream_import_resolves() {
    assert!(
        invoke_result("wasi:cli/stdin", "read-via-stream", vec![]).is_ok(),
        "wasi:cli/stdin.read-via-stream should be covered by the CLI category"
    );
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

#[test]
fn proposal_cli_exit_import_is_registered() {
    assert!(
        has_import("wasi:cli/exit", "exit"),
        "wasi:cli/exit.exit should be covered by the CLI category"
    );
}

#[test]
fn proposal_cli_terminal_stdin_import_resolves() {
    assert!(
        invoke_result("wasi:cli/terminal-stdin", "get-terminal-stdin", vec![]).is_ok(),
        "wasi:cli/terminal-stdin.get-terminal-stdin should be covered by the CLI category"
    );
}

#[test]
fn proposal_cli_terminal_stdout_import_resolves() {
    assert!(
        invoke_result("wasi:cli/terminal-stdout", "get-terminal-stdout", vec![]).is_ok(),
        "wasi:cli/terminal-stdout.get-terminal-stdout should be covered by the CLI category"
    );
}

#[test]
fn proposal_cli_terminal_stderr_import_resolves() {
    assert!(
        invoke_result("wasi:cli/terminal-stderr", "get-terminal-stderr", vec![]).is_ok(),
        "wasi:cli/terminal-stderr.get-terminal-stderr should be covered by the CLI category"
    );
}