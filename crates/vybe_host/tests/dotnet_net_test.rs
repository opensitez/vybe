use std::sync::Arc;
use vybe_bytecode::value::{ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_import(module: &str, name: &str, pre_stack: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<dotnet-net-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = pre_stack.len() as u8;
    for value in pre_stack {
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

fn object_prop(value: &Value, key: &str) -> Value {
    match value {
        Value::Object(obj) => obj.lock().unwrap().properties.get(key).cloned().unwrap_or(Value::Null),
        _ => Value::Null,
    }
}

fn array_len(value: &Value) -> usize {
    match value {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elements) => elements.len(),
                _ => 0,
            }
        }
        _ => 0,
    }
}

#[test]
fn dotnet_net_dns_host_addresses_returns_array() {
    let result = call_import("dotnet:net", "dnsGetHostAddresses", vec![Value::String(Arc::from("localhost"))]);
    assert!(array_len(&result) > 0, "localhost should resolve to at least one address");
}

#[test]
fn dotnet_net_dns_host_entry_returns_hostname_and_addresses() {
    let result = call_import("dotnet:net", "dnsGetHostEntry", vec![Value::String(Arc::from("localhost"))]);
    assert_eq!(object_prop(&result, "hostname").as_str(), "localhost");
    assert!(array_len(&object_prop(&result, "addresslist")) > 0, "host entry should include addresses");
}

#[test]
fn dotnet_net_dns_host_name_returns_string() {
    let result = call_import("dotnet:net", "dnsGetHostName", vec![]);
    assert!(!result.as_str().is_empty(), "machine name should not be empty");
}