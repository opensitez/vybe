//! Go `encoding/*` adapter helpers that need byte-level loops.

use vybe_compiler::primitives::{collections, convert, loops};
use vybe_compiler::primitives::instructions::host;
use vybe_runtime::{Chunk, opcode::Op};

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.hex.EncodeToString" if argc == 1 => emit_hex_encode_to_string(chunks, current, line),
        "go.hex.DecodeStringBytes" if argc == 1 => {
            emit_hex_decode_string_bytes(chunks, current, line)
        }
        _ => return false,
    }
    true
}

fn emit_hex_encode_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    // Go nil byte slices encode as empty; preserve all non-null inputs.
    collections::emit_nil_to_empty_array(chunks, current, line);
    let base=chunks[current].alloc_scratch(6);
    let bytes=base;let parts=base+1;let i=base+2;let len=base+3;let code=base+4;let table=base+5;
    set(&mut chunks[current],bytes,line);
    collections::emit_array_new(chunks,current,0,line);set(&mut chunks[current],parts,line);
    get(&mut chunks[current],bytes,line);collections::emit_len(chunks,current,line);
    convert::emit_to_int(&mut chunks[current],line);set(&mut chunks[current],len,line);
    chunks[current].emit_string_const("0123456789abcdef",line);set(&mut chunks[current],table,line);
    chunks[current].emit_i32_const(0,line);set(&mut chunks[current],i,line);
    let state=loops::emit_loop_start(chunks,current,line);
    get(&mut chunks[current],i,line);get(&mut chunks[current],len,line);
    chunks[current].emit_op(Op::I32_LT_S,line);loops::emit_loop_cond_from_i32(chunks,current,line);

    get(&mut chunks[current],bytes,line);get(&mut chunks[current],i,line);
    collections::emit_get(chunks,current,line);
    // Retain the exact old byte-to-UTF16-unit coercion, including baseline
    // malformed byte-width behavior, without assembling a whole binary string.
    host::emit(&mut chunks[current],"wasm:js-string","fromCharCode",1,line);
    chunks[current].emit_i32_const(0,line);
    host::emit(&mut chunks[current],"wasm:js-string","charCodeAt",2,line);
    set(&mut chunks[current],code,line);
    for high in [true,false] {
        get(&mut chunks[current],parts,line);get(&mut chunks[current],table,line);
        get(&mut chunks[current],code,line);
        chunks[current].emit_i32_const(if high {4} else {15},line);
        chunks[current].emit_op(if high {Op::I32_SHR_U} else {Op::I32_AND},line);
        host::emit(&mut chunks[current],"ecma:string","charAt",2,line);
        collections::emit_push(chunks,current,line);chunks[current].emit_op(Op::DROP,line);
    }
    get(&mut chunks[current],i,line);chunks[current].emit_i32_const(1,line);
    chunks[current].emit_op(Op::I32_ADD,line);set(&mut chunks[current],i,line);
    loops::emit_loop_end(chunks,current,state,line);
    get(&mut chunks[current],parts,line);chunks[current].emit_string_const("",line);
    host::emit(&mut chunks[current],"ecma:array","join",2,line);
}

fn emit_hex_decode_string_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::string_encoding::emit_hex2bytes(chunks, current, 1, line);
}

fn get(c: &mut Chunk, slot: u16, line: u32) { c.emit_op_u16(Op::LOCAL_GET,slot,line); }
fn set(c: &mut Chunk, slot: u16, line: u32) { c.emit_op_u16(Op::LOCAL_SET,slot,line); }

// The compiler's exact-parser/producer Owner is the sole reuse authority.
use vybe_compiler::primitives::helper_templates::{Contract, Definition};
const HEX_BODY: &str = "\0go.owned.hex.encode.v1";
fn hex_context_key(chunks: &[Chunk], _: u32) -> Vec<u8> {
    vec![match vybe_compiler::primitives::class_context::module_receiver_abi(chunks) {
        vybe_runtime::chunk::ReceiverAbi::Ambient => 0,
        vybe_runtime::chunk::ReceiverAbi::Parameter => 1,
    }]
}
fn unused_hex_factory() -> Chunk { unreachable!("contextual recipe only") }
fn unused_hex_caller(_: &mut Chunk, _: usize, _: u32) { unreachable!("contextual recipe only") }
fn build_hex_body(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let index=chunks.len();
    let mut body=Chunk::new(HEX_BODY);
    body.arity=1;body.param_count=1;body.local_count=1;
    body.takes_receiver=false;body.direct_entry_has_captures=Some(false);
    body.module_receiver_abi=vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    body.emit_op_u16(Op::LOCAL_GET,0,line);
    chunks.push(body);
    emit_hex_encode_to_string(chunks,index,line);
    chunks[index].emit_op(Op::RETURN,line);
    index
}
fn invoke_hex_body(chunks:&mut Vec<Chunk>,current:usize,target:usize,line:u32) {
    let abi=vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    let valid=chunks.get(target).is_some_and(|c|
        c.name==HEX_BODY && c.arity==1 && c.param_count==1 && c.result_arity==1
        && c.capture_count==0 && c.direct_entry_has_captures==Some(false)
        && !c.takes_receiver && !c.is_method && !c.is_async && !c.is_generator
        && c.module_receiver_abi==abi && c.func_sig.is_none()
        && c.native_func_type.is_none() && c.local_count>=1);
    if !valid || current==target {emit_hex_encode_to_string(chunks,current,line);return;}
    let (caller,callee)=if current<target {
        let(a,b)=chunks.split_at_mut(target);(&mut a[current],&b[0])
    }else{let(a,b)=chunks.split_at_mut(current);(&mut b[0],&a[target])};
    caller.emit_direct_defined_call(target as u32,callee,false,line)
        .expect("owned plain hex body has no captures or continuation");
}
pub static HELPER_DEFINITIONS: &[Definition] = &[Definition {
    contract:Contract::ContextualAppendOnlyUnit{key:hex_context_key,build:build_hex_body,invoke:invoke_hex_body},
    operation:"go.hex.EncodeToString",namespace:"go",version:1,
    parser:crate::parse,producer:super::dispatch::dispatch,arity:1,
    helper_name:HEX_BODY,factory:unused_hex_factory,caller:unused_hex_caller,
}];
