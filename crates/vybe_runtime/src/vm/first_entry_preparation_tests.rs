use super::*;
use crate::{PreparedLinkedCode, Value};
fn module()->Vec<Chunk>{
    let mut script=Chunk::new("<first-entry-script>");
    script.globals=Arc::new(vec!["answer".into()].into());
    script.emit_op_u16(Op::REF_FUNC,1,7);script.emit(0,7);script.emit_op(Op::RETURN,7);
    let mut outer=Chunk::new("first-entry-outer");
    outer.globals=script.globals.clone();
    // NOP makes instruction starts deliberately unaligned after REF_FUNC.
    outer.emit_op_u16(Op::REF_FUNC,2,8);outer.emit(0,8);outer.emit_op(Op::RETURN,8);
    let mut leaf=Chunk::new("first-entry-leaf");leaf.globals=script.globals.clone();
    leaf.emit_block_typed(9,1);leaf.emit_op_u32(Op::GLOBAL_GET,0,9);leaf.emit_op(Op::END,9);leaf.emit_op(Op::RETURN,9);
    let mut unused=Chunk::new("unused-large-body");
    for _ in 0..2048{unused.emit_f64_const(1.25,10);unused.emit_op(Op::DROP,10);}
    unused.emit_ref_null(crate::opcode::heaptype::HT_EXTERN,10);unused.emit_op(Op::RETURN,10);
    vec![script,outer,leaf,unused]
}
fn vm()->VM{
    let mut vm=VM::new();vm.jit_enabled=false;
    vm.set_global("prior",Value::F64(11.));vm.set_global("answer",Value::F64(41.));
    let mut prior=Chunk::new("prior-unit");prior.emit_f64_const(0.,1);prior.emit_op(Op::RETURN,1);
    vm.run_linked_nested(vec![prior],vec![]).unwrap();vm
}
#[test]
fn first_entry_sparse_preparation_keeps_complete_code_and_reduces_metadata(){
    let raw=module();let before=format!("{raw:?}");
    let eager=PreparedLinkedCode::prepare(raw.clone(),usize::MAX).ok().unwrap();
    let lazy=PreparedLinkedCode::prepare_for_first_entry(raw,usize::MAX).ok().unwrap();
    assert_eq!(format!("{:?}",lazy.chunks()),before);
    assert!(lazy.metadata_bytes()*8<eager.metadata_bytes());
    let(_,rows)=lazy.into_parts();let rows=rows.unwrap();assert!(!rows[0].deferred_entry);
    assert!(rows[1..].iter().all(|r|r.deferred_entry&&r.ops.is_empty()&&r.operands.is_empty()&&r.wide.is_empty()&&r.blocks.is_empty()&&r.br_tables.is_empty()));
    assert_eq!(rows[1].refs.len(),1);assert_eq!(rows[2].globals.len(),1);
}
#[test]
fn first_entry_nested_ref_and_global_bases_match_eager_and_prepare_only_once(){
    let mut eager=vm();let mut lazy=vm();
    let normal=PreparedLinkedCode::prepare(module(),usize::MAX).ok().unwrap();
    let sparse=PreparedLinkedCode::prepare_for_first_entry(module(),usize::MAX).ok().unwrap();
    let a=eager.run_linked_nested_prepared(normal,vec![]).unwrap();
    let b=lazy.run_linked_nested_prepared(sparse,vec![]).unwrap();
    assert_eq!(lazy.pending_entry_decode.len(),3);
    for i in 2..5{assert!(lazy.decoded_ops[i].is_empty());assert!(lazy.block_targets[i].indices.is_empty());}
    let af=eager.invoke(&a,&[]).unwrap();let bf=lazy.invoke(&b,&[]).unwrap();
    assert_eq!(lazy.pending_entry_decode.len(),2);assert_eq!(lazy.decoded_operands64[2],eager.decoded_operands64[2]);
    assert_eq!(eager.invoke(&af,&[]).unwrap(),Value::F64(41.));assert_eq!(lazy.invoke(&bf,&[]).unwrap(),Value::F64(41.));
    assert_eq!(lazy.pending_entry_decode.len(),1);assert!(lazy.pending_entry_decode.contains_key(&4));
    for i in 1..4{
        assert_eq!(lazy.decoded_ops[i],eager.decoded_ops[i]);assert_eq!(lazy.decoded_operands[i],eager.decoded_operands[i]);assert_eq!(lazy.decoded_operands64[i],eager.decoded_operands64[i]);
        assert_eq!(format!("{:?}",lazy.block_targets[i]),format!("{:?}",eager.block_targets[i]));
    }
    let pointer=lazy.decoded_ops[2].as_ptr();let wide=lazy.decoded_operands64[2].clone();
    let again=lazy.invoke(&b,&[]).unwrap();assert_eq!(lazy.invoke(&again,&[]).unwrap(),Value::F64(41.));
    assert_eq!(pointer,lazy.decoded_ops[2].as_ptr());assert_eq!(wide,lazy.decoded_operands64[2]);assert!(lazy.decoded_ops[4].is_empty());
}
#[test]
fn first_entry_reset_drops_pending_units_and_reused_chunk_slots_get_fresh_base(){
    let mut vm=vm();let snapshot=vm.snapshot();
    let old=vm.run_linked_nested_prepared(PreparedLinkedCode::prepare_for_first_entry(module(),usize::MAX).ok().unwrap(),vec![]).unwrap();
    assert_eq!(vm.pending_entry_decode.len(),3);vm.reset_to(&snapshot);assert!(vm.pending_entry_decode.is_empty());
    assert!(vm.invoke(&old,&[]).is_err());
    let new=vm.run_linked_nested_prepared(PreparedLinkedCode::prepare_for_first_entry(module(),usize::MAX).ok().unwrap(),vec![]).unwrap();
    let leaf=vm.invoke(&new,&[]).unwrap();assert_eq!(vm.invoke(&leaf,&[]).unwrap(),Value::F64(41.));
}
#[test]
fn first_entry_unused_invalid_body_still_fails_at_load(){
    let mut code=module();code[3].code=vec![255,255,255,255];code[3].lines.clear();
    let mut eager=vm();let mut lazy=vm();let before=lazy.chunks.len();
    let a=eager.run_linked_nested_prepared(PreparedLinkedCode::prepare(code.clone(),usize::MAX).ok().unwrap(),vec![]).unwrap_err();
    let b=lazy.run_linked_nested_prepared(PreparedLinkedCode::prepare_for_first_entry(code,usize::MAX).ok().unwrap(),vec![]).unwrap_err();
    assert_eq!(format!("{a:?}"),format!("{b:?}"));assert_eq!(lazy.chunks.len(),before);assert!(lazy.pending_entry_decode.is_empty());
}
#[test]
fn first_entry_branch_table_cache_is_built_only_when_body_enters(){
    let mut code=module();let leaf=&mut code[2];leaf.code.clear();leaf.lines.clear();
    leaf.emit_block(9);leaf.emit_i32_const(0,9);leaf.emit_br_table(&[0,0,0],0,9);leaf.emit_op(Op::END,9);
    leaf.emit_op_u32(Op::GLOBAL_GET,0,9);leaf.emit_op(Op::RETURN,9);
    let mut eager=vm();let mut lazy=vm();
    let a=eager.run_linked_nested_prepared(PreparedLinkedCode::prepare(code.clone(),usize::MAX).ok().unwrap(),vec![]).unwrap();
    let b=lazy.run_linked_nested_prepared(PreparedLinkedCode::prepare_for_first_entry(code,usize::MAX).ok().unwrap(),vec![]).unwrap();
    assert!(lazy.br_table_targets[3].is_empty());let a=eager.invoke(&a,&[]).unwrap();let b=lazy.invoke(&b,&[]).unwrap();
    assert_eq!(eager.invoke(&a,&[]).unwrap(),Value::F64(41.));assert_eq!(lazy.invoke(&b,&[]).unwrap(),Value::F64(41.));
    assert!(lazy.br_table_targets[3].iter().any(Option::is_some));
    assert_eq!(format!("{:?}",lazy.br_table_targets[3]),format!("{:?}",eager.br_table_targets[3]));
}
#[test]
fn first_entry_decode_installs_original_native_targets_before_fusion(){
    let mut body=Chunk::new("first-entry-native-target");body.arity=1;body.local_count=1;
    body.emit_op_u16(Op::LOCAL_GET,0,1);let call=body.code.len();body.emit_call(0,1,1);body.emit_op_u16(Op::LOCAL_SET,0,1);body.emit_op_u16(Op::LOCAL_GET,0,1);body.emit_op(Op::RETURN,1);
    let target=NativeCallSiteTarget::JsString(JsStringBuiltin::Length);let mut expected=vec![None;body.code.len().div_ceil(4)];expected[call/4]=Some(target);
    let caches=VM::build_execution_fast_paths_for_chunk(&body,Some(&expected));
    let mut vm=VM::new();vm.chunks.push(body);vm.decoded_ops.push(Vec::new());vm.decoded_operands.push(Vec::new());vm.decoded_operands64.push(Vec::new());vm.br_table_targets.push(Vec::new());vm.block_targets.push(Arc::new(BlockTargetCache::default()));vm.native_call_sites.push(Vec::new());vm.local_fast_paths.push(Vec::new());vm.native_result_fast_paths.push(Vec::new());
    vm.pending_entry_decode.insert(0,0);vm.pending_native_call_sites.insert(0,vec![(call,target)]);
    // Another module's restored transient context must not retarget this site.
    vm.resolved_import_table=vec![ResolvedCallTarget::JsBoolean(JsBooleanBuiltin::Cast)];vm.ensure_execution_fast_paths(0);
    assert!(vm.pending_entry_decode.is_empty());assert!(vm.pending_native_call_sites.is_empty());
    assert_eq!(vm.native_call_sites[0],expected);assert_eq!(vm.local_fast_paths[0],caches.0);assert_eq!(vm.native_result_fast_paths[0],caches.1);
}

#[test]
fn first_entry_reload_unentered_body_preserves_other_pending_module_bases() {
    let mut eager=vm(); let mut lazy=vm();
    let normal=PreparedLinkedCode::prepare(module(),usize::MAX).ok().unwrap();
    let sparse=PreparedLinkedCode::prepare_for_first_entry(module(),usize::MAX).ok().unwrap();
    let a=eager.run_linked_nested_prepared(normal,vec![]).unwrap();
    let b=lazy.run_linked_nested_prepared(sparse,vec![]).unwrap();
    assert!(lazy.pending_entry_decode.contains_key(&2));
    // Reload accepts the complete installed namespace, hence replacement
    // REF_FUNC operands are absolute, unlike an incoming appended unit.
    let mut replacement=lazy.chunks.clone();
    let mut body=Chunk::new("first-entry-outer");
    body.globals=replacement[2].globals.clone();
    body.emit_op_u16(Op::REF_FUNC,3,8); body.emit(0,8); body.emit_op(Op::RETURN,8);
    replacement[2]=body;
    assert_eq!(eager.apply_reload(replacement.clone()).unwrap(),lazy.apply_reload(replacement).unwrap());
    assert!(!lazy.pending_entry_decode.contains_key(&2));
    assert!(lazy.pending_entry_decode.contains_key(&3));
    assert!(lazy.pending_entry_decode.contains_key(&4));
    let af=eager.invoke(&a,&[]).unwrap(); let bf=lazy.invoke(&b,&[]).unwrap();
    assert_eq!(eager.invoke(&af,&[]).unwrap(),Value::F64(41.));
    assert_eq!(lazy.invoke(&bf,&[]).unwrap(),Value::F64(41.));
    assert_eq!(lazy.decoded_operands64[2],eager.decoded_operands64[2]);
    assert!(lazy.pending_entry_decode.contains_key(&4));
}

#[test]
fn first_entry_reload_live_frame_keeps_old_body_and_updates_future_calls() {
    let mut vm=vm();
    let retained=vm.run_linked_nested_prepared(PreparedLinkedCode::prepare_for_first_entry(module(),usize::MAX).ok().unwrap(),vec![]).unwrap();
    vm.stack.push(retained.clone()); vm.call_value(0).unwrap();
    assert_eq!(vm.frames.last().unwrap().chunk_index,2);
    let mut replacement=vm.chunks.clone();
    let mut body=Chunk::new("first-entry-outer"); body.emit_f64_const(99.,8);body.emit_op(Op::RETURN,8);
    replacement[2]=body;
    let report=vm.apply_reload(replacement).unwrap();
    assert!(report.contains("live frame kept on old body"));
    let relocated=vm.frames.last().unwrap().chunk_index;
    assert!(relocated>=5); assert!(!vm.pending_entry_decode.contains_key(&relocated));
    let leaf=vm.execute().unwrap();assert_eq!(vm.invoke(&leaf,&[]).unwrap(),Value::F64(41.));
    assert_eq!(vm.invoke(&retained,&[]).unwrap(),Value::F64(99.));
    assert!(vm.pending_entry_decode.contains_key(&4));
}
