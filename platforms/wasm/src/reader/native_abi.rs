use super::read_wasm;
use crate::encoding::*;
use vybe_runtime::chunk::ReceiverAbi;

const SOURCE_ABI: &[u8] = b"\x01\x01\x04last\x02\x03\x01\x01";

fn module(indirect: bool, extra_local: bool, abi: Option<&[u8]>) -> Vec<u8> {
    let mut bytes = Vec::from(WASM_MAGIC);
    bytes.extend_from_slice(&WASM_VERSION);
    write_section(&mut bytes, SECTION_TYPE,
        &[2, TYPE_FUNC, 3, TYPE_EXTERNREF, TYPE_EXTERNREF, TYPE_I32, 1, TYPE_I32,
            TYPE_FUNC, 0, 0]);
    write_section(&mut bytes, SECTION_FUNCTION, &[2, 0, 1]);
    let mut exports = vec![1];
    write_name(&mut exports, "_start");
    exports.extend_from_slice(&[0, 1]);
    write_section(&mut bytes, SECTION_EXPORT, &exports);
    write_section(&mut bytes, 9, &[1, 3, 0, 1, 0]);

    // last(null, null, 42) must return the third native parameter, not a local.
    let last: &[u8] = if extra_local {
        &[1, 1, TYPE_I32, 0x20, 2, 0x21, 3, 0x20, 3, 0x0b]
    } else {
        &[0, 0x20, 2, 0x0b]
    };
    let mut entry = vec![0, 0xd0, TYPE_EXTERNREF, 0xd0, TYPE_EXTERNREF, 0x41, 42];
    if indirect {
        entry.extend_from_slice(&[0xd2, 0, 0x14, 0]);
    } else {
        entry.extend_from_slice(&[0x10, 0]);
    }
    entry.extend_from_slice(&[0x41, 42, 0x47, 0x04, 0x40, 0, 0x0b, 0x0b]);
    let mut code = vec![2];
    write_leb128_u32(&mut code, last.len() as u32);
    code.extend_from_slice(last);
    write_leb128_u32(&mut code, entry.len() as u32);
    code.extend_from_slice(&entry);
    write_section(&mut bytes, SECTION_CODE, &code);
    if let Some(abi) = abi {
        let mut custom = Vec::new();
        write_name(&mut custom, "vybe.abi");
        custom.extend_from_slice(abi);
        write_section(&mut bytes, SECTION_CUSTOM, &custom);
    }
    // Standard names remain usable without supplying execution semantics.
    let mut names = Vec::new();
    write_name(&mut names, "name");
    names.extend_from_slice(b"\x01\x07\x01\x00\x04last\x04\x0c\x01\x00\x09signature");
    write_section(&mut bytes, SECTION_CUSTOM, &names);
    bytes
}

fn check(indirect: bool, extra_local: bool, abi: Option<&[u8]>) {
    let bytes = module(indirect, extra_local, abi);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).expect("custom payloads do not change native validity");
    let chunks = read_wasm(&bytes).unwrap();
    let last = chunks.iter().find(|chunk| chunk.name == "last").unwrap();
    assert_eq!(last.arity, 3);
    assert_eq!(last.param_count, 3);
    assert_eq!(last.local_count, if extra_local { 4 } else { 3 });
    assert!(!last.takes_receiver);
    assert_eq!(last.module_receiver_abi, ReceiverAbi::Ambient);
    assert_eq!(last.func_sig.as_ref().unwrap(),
        &("externref,externref,i32".into(), "i32".into()));
    for linked in [false, true] {
        let mut vm = vybe_runtime::VM::new();
        let result = if linked {
            vm.run_linked(chunks.clone(), vec![])
        } else {
            vm.run(chunks.clone())
        };
        result.unwrap_or_else(|error| panic!("indirect={indirect} linked={linked}: {error}"));
        assert_eq!(vm.pending_exit_code, 0);
    }
}

#[test]
fn source_abi_cannot_truncate_native_call_ref_arguments() {
    assert_eq!(module(true, false, Some(SOURCE_ABI)).len(), 125);
    check(true, false, Some(SOURCE_ABI));
    check(true, false, None);
}

#[test]
fn source_abi_cannot_change_native_direct_arguments_or_local_layout() {
    for indirect in [false, true] {
        check(indirect, true, Some(SOURCE_ABI));
        check(indirect, true, Some(b"\x01\x01\x04last\x00\xc8\x01\x01\x01"));
        check(indirect, true, None);
    }
}

#[test]
fn malformed_source_abi_payload_is_opaque_to_native_execution() {
    for payload in [&b""[..], &b"\xff"[..], &b"\x01\x01\x80"[..],
        &b"\x01\x01\x01\xff\x00\x00\x01\x01"[..]]
    {
        for indirect in [false, true] {
            check(indirect, false, Some(payload));
        }
    }
}

#[test]
fn opaque_abi_does_not_waive_native_validation() {
    let mut bytes = module(true, false, Some(b"\xff"));
    bytes[8] = SECTION_CODE;
    assert!(super::read_wasm_classified(&bytes).is_err());
}

#[test]
fn custom_ir_branch_retains_its_source_abi() {
    let mut script = vybe_runtime::Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(vybe_runtime::Op::RETURN, 0);
    let mut function = vybe_runtime::Chunk::new("source_receiver");
    function.arity = 2;
    function.param_count = 2;
    function.local_count = 5;
    function.takes_receiver = true;
    function.module_receiver_abi = ReceiverAbi::Parameter;
    function.emit_op_u16(vybe_runtime::Op::LOCAL_GET, 1, 0);
    function.emit_op(vybe_runtime::Op::RETURN, 0);
    let bytes = crate::writer::write_wasm_with_options(&[script, function],
        crate::writer::WasmWriteOptions {
            include_vybe_metadata: true,
            include_name_section: true,
        });
    let chunks = read_wasm(&bytes).unwrap();
    let function = chunks.iter().find(|chunk| chunk.name == "source_receiver").unwrap();
    assert_eq!(function.arity, 2);
    assert_eq!(function.local_count, 5);
    assert!(function.takes_receiver);
    assert_eq!(function.module_receiver_abi, ReceiverAbi::Parameter);
}
