//! One validated source-unit preparation, incremental ordinary emission, one
//! finalization. This is not pre-EOF execution: provisional chunks stay private.
use std::sync::Arc;
use vybe_compiler::Compiler;
use vybe_runtime::{Chunk, PreparedLinkedCode, Value, VM};
#[path = "../examples/emission_import_interning/canonical.rs"]
mod metadata;

fn compiler() -> Compiler {
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    Compiler::with_profile(vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap())
        .with_optimizations(vybe_runtime::registry::compiler_optimizations(language))
        .with_helper_templates(vybe_language_php::parse,
            vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,
            Arc::new(vybe_compiler::primitives::helper_templates::Pool::default()))
        .with_argument_emitters(vybe_language_php::parse,
            vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS)
        .with_run_module_end_hooks(false)
}
fn same(a: &[Chunk], b: &[Chunk]) {
    assert_eq!(a.len(), b.len());
    for (a,b) in a.iter().zip(b) {
        assert_eq!(a.code,b.code); assert_eq!(a.lines,b.lines);
        assert_eq!(metadata::canonical(&format!("{a:?}")),metadata::canonical(&format!("{b:?}")));
    }
    let write=|c|vybe_platform_wasm::write_wasm_with_options(c,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata:false, include_name_section:true });
    let a=write(a);let b=write(b);assert_eq!(a,b);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&a).unwrap();
}
fn run(chunks: Vec<Chunk>, prepared: bool, expected: i32) -> String {
    let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let resolved=chunks[0].imports.iter().map(|i|vm.resolve_import_target(&i.module,&i.name)
        .or_else(|e|if matches!(i.module.as_str(),"env"|"*") {
            Ok(vybe_runtime::ImportTarget::StdlibRedirect(i.name.clone()))
        } else { Err(e) })).collect::<Result<Vec<_>,_>>().unwrap();
    let value=if prepared {
        vm.run_linked_prepared(PreparedLinkedCode::prepare(chunks,64*1024*1024).ok().unwrap(),resolved)
    } else { vm.run_linked(chunks,resolved) }.unwrap();
    let number=match &value {Value::I32(n)=>*n as f64,Value::I64(n)=>*n as f64,Value::F64(n)=>*n,
        other=>panic!("expected numeric {expected}, got {other:?}")};
    assert_eq!(number,expected as f64);
    format!("{value:?}")
}
fn check(source: &str, expected: i32) {
    let ast=vybe_language_php::parse(source).unwrap();
    let before=format!("{ast:?}");
    let batch=compiler().compile_with_imports(&ast).unwrap();
    let mut steps=compiler().begin_emission(&ast).unwrap();
    let mut body_seen=false;let mut count=0;
    while let Some(progress)=steps.emit_next().unwrap() {
        if !progress.function_declaration {body_seen=true;}
        assert!(!progress.function_declaration || !body_seen,"hoisted functions must precede body");
        let pending=steps.pending_chunks().unwrap();
        assert!(progress.appended_chunks.end<=pending.len());
        assert!(progress.entry_bytes.end<=pending[0].code.len());
        count+=1;
        // Scope must restore between steps; callers may perform other emission.
        let _unrelated=vybe_runtime::source_lines::LineEmissionScope::enter(false);
    }
    assert!(count>0);assert!(steps.emit_next().unwrap().is_none());
    let incremental=steps.finish().unwrap();
    assert_eq!(before,format!("{ast:?}"),"borrowed AST remains unchanged");
    let owned=compiler().begin_emission_owned(ast).unwrap().finish().unwrap();
    assert_eq!(format!("{:?}",batch.host_imports),format!("{:?}",incremental.host_imports));
    assert_eq!(batch.component_source,incremental.component_source);
    assert_eq!(batch.app_shell,incremental.app_shell);
    same(&batch.chunks,&incremental.chunks);same(&batch.chunks,&owned.chunks);
    let oracle=run(batch.chunks.clone(),false,expected);
    assert_eq!(run(batch.chunks,true,expected),oracle);
    assert_eq!(run(incremental.chunks.clone(),false,expected),oracle);
    assert_eq!(run(incremental.chunks,true,expected),oracle);
    assert_eq!(run(owned.chunks,false,expected),oracle);
}
#[test]
fn late_declaration_and_rest_parameter_manifest_are_visible_before_first_emit() {
    check("<?php $r=f(2,3,4); function f($x,...$rest){return $x+$rest[0]+$rest[1];} return $r;",9);
}
#[test]
fn address_promotion_and_nested_control_flow_share_one_compiler_state() {
    check("<?php function setv(&$x){$x=7;} $v=1;setv($v);$n=0;for($i=0;$i<3;$i++){$n++;}if($n===3){return $v;}return 0;",7);
}
#[test]
fn namespace_inheritance_properties_and_closures_use_existing_emitters() {
    check("<?php namespace StreamTest; class Base {public $x=4; public function get(){return $this->x;}} class Child extends Base {public function value(){return $this->get()+3;}} $c=new Child();$f=function($o){return $o->value();};return $f($c);",7);
}
#[test]
fn abort_discards_unpublished_unit_and_does_not_poison_next_unit() {
    let ast=vybe_language_php::parse("<?php function unused(){return 99;} $x=1;return $x;").unwrap();
    let mut session=compiler().begin_emission(&ast).unwrap();
    assert!(session.emit_next().unwrap().unwrap().function_declaration);
    assert!(session.pending_chunks().unwrap().len()>1);
    session.abort();
    check("<?php return 7;",7);
}
#[test]
fn common_ast_emission_failure_poison_discards_partial_metadata() {
    let mut ast=vybe_language_php::parse("<?php $x=1;return $x;").unwrap();
    // This normal common-AST producer mutates data_segments before rejecting its
    // invalid offset. Rewinding only code length would leave corrupt metadata.
    ast.body.push(vybe_ast::Statement::new(vybe_ast::StmtKind::DataSegment {
        memory_index:0, offset:Some(vybe_ast::Expression::new(vybe_ast::ExprKind::Ident("not_a_constant".into()))),
        bytes:vec![1,2,3],
    }));
    let mut session=compiler().begin_emission_owned(ast).unwrap();
    let error=loop {match session.emit_next() {
        Ok(Some(_))=>{},Ok(None)=>panic!("invalid producer must fail"),Err(error)=>break error,
    }};
    assert_eq!(error,"data segment offset must be a constant integer expression");
    assert!(session.pending_chunks().is_err());
    assert_eq!(session.emit_next().unwrap_err(),"emission session is poisoned");
    assert_eq!(session.finish().unwrap_err(),"emission session is poisoned");
}

fn check_statement_stream(mut ast:vybe_ast::Module, expected:i32) {
    let batch=compiler().compile_with_imports(&ast).unwrap();
    let mut stream=compiler().begin_statement_emission(&ast).unwrap();
    let count=ast.body.len();assert_eq!(stream.expected_statement_count(),count);
    // Move producer-owned statements; the compiler stream has no lifetime tied
    // to Module and cannot retain its body borrow. No source AST clone here.
    let body=std::mem::take(&mut ast.body);drop(ast);
    for(index,statement)in body.into_iter().enumerate(){
        let function=matches!(statement.kind,vybe_ast::StmtKind::FunctionDecl{..});
        let before=stream.pending_chunks().unwrap()[0].code.len();
        let progress=stream.emit_statement(index,statement).unwrap();
        assert_eq!(progress.source_index,index);
        if function{assert!(!progress.emitted,"manifest functions must not emit twice");}
        assert_eq!(progress.entry_bytes.start,before);
        assert_eq!(stream.consumed_statement_count(),index+1);
    }
    let result=stream.finish().unwrap();
    same(&batch.chunks,&result.chunks);
    assert_eq!(format!("{:?}",batch.host_imports),format!("{:?}",result.host_imports));
    assert_eq!(batch.component_source,result.component_source);
    assert_eq!(batch.app_shell,result.app_shell);
    let oracle=run(batch.chunks.clone(),false,expected);
    assert_eq!(run(batch.chunks,true,expected),oracle);
    assert_eq!(run(result.chunks.clone(),false,expected),oracle);
    assert_eq!(run(result.chunks,true,expected),oracle);
}
#[test]
fn producer_statements_emit_native_chunks_with_forward_function_and_rest_facts() {
    check_statement_stream(vybe_language_php::parse("<?php $r=f(2,3,4);function f($x,...$rest){return $x+$rest[0]+$rest[1];}return $r;").unwrap(),9);
}
#[test]
fn producer_statements_preserve_namespaces_classes_closures_and_pointer_places() {
    check_statement_stream(vybe_language_php::parse("<?php namespace StreamTest; class Base {public $x=4;public function get(){return $this->x;}}class Child extends Base{public function value(){return $this->get()+3;}}$c=new Child();$f=function($o){return $o->value();};return $f($c);").unwrap(),7);
    check_statement_stream(vybe_language_php::parse("<?php $a=3;$b=&$a;$b=8;return $a;").unwrap(),8);
}
#[test]
fn statement_manifest_rejects_missing_reordered_and_changed_producers() {
    let mut ast=vybe_language_php::parse("<?php $a=1;return 7;").unwrap();
    let incomplete=compiler().begin_statement_emission(&ast).unwrap();
    assert!(incomplete.finish().unwrap_err().contains("producer incomplete"));
    let mut reordered=compiler().begin_statement_emission(&ast).unwrap();
    assert!(reordered.emit_statement(1,ast.body[0].clone()).unwrap_err().contains("out of order"));
    assert!(reordered.pending_chunks().is_err());
    let mut changed=compiler().begin_statement_emission(&ast).unwrap();
    let mut statement=ast.body.remove(0);
    statement.walk_exprs_mut(&mut|expression|if let vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(value))=&mut expression.kind{*value+=1;});
    assert!(changed.emit_statement(0,statement).unwrap_err().contains("differs from sealed normalized manifest"));
    assert!(changed.pending_chunks().is_err());
}
#[test]
fn statement_manifest_keeps_legacy_partial_class_group_order() {
    let mut ast=vybe_language_php::parse("<?php class Merge {public function left(){return 3;}public function right(){return 4;}}$m=new Merge();return $m->left()+$m->right();").unwrap();
    let index=ast.body.iter().position(|s|matches!(&s.kind,vybe_ast::StmtKind::ClassDecl{name,..}if name=="Merge")).unwrap();
    let mut second=ast.body[index].clone();
    let is_right=|member:&vybe_ast::ClassMember|matches!(member,
        vybe_ast::ClassMember::Method(method) if matches!(&method.kind,
            vybe_ast::StmtKind::FunctionDecl{name,..}if name=="right"));
    if let vybe_ast::StmtKind::ClassDecl{members,modifiers,..}=&mut second.kind {
        members.retain(is_right);assert_eq!(members.len(),1);modifiers.is_partial=true;
    }
    if let vybe_ast::StmtKind::ClassDecl{members,modifiers,..}=&mut ast.body[index].kind {
        members.retain(|member|!is_right(member));modifiers.is_partial=true;
    }
    // Keep original parser-normalized names and owner metadata; only split the
    // existing valid declaration into two ordinary commonAST partial headers.
    ast.body.insert(index+1,second);check_statement_stream(ast,7);
}
#[test]
fn grammar_failure_can_abort_incremental_instructions_without_finalization() {
    let mut ast=vybe_language_php::parse("<?php $a=1;$a+=2;return $a;").unwrap();
    let mut stream=compiler().begin_statement_emission(&ast).unwrap();
    let statement=ast.body.remove(0);
    let before=stream.pending_chunks().unwrap()[0].code.len();
    let progress=stream.emit_statement(0,statement).unwrap();
    assert!(progress.emitted);assert!(stream.pending_chunks().unwrap()[0].code.len()>before);
    // The selected grammar entry subsequently fails: discard pending code.
    stream.abort();
    check_statement_stream(vybe_language_php::parse("<?php return 7;").unwrap(),7);
}

fn service_vm() -> VM {
    vybe_language_php::register();
    let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    vm
}
fn numeric_service_value(value: Value, expected: i32) {
    let number=match value {Value::I32(n)=>n as f64,Value::I64(n)=>n as f64,
        Value::F64(n)=>n,other=>panic!("expected numeric {expected}, got {other:?}")};
    assert_eq!(number,expected as f64);
}
#[test]
fn registered_php_public_compiler_service_uses_complete_source_unit() {
    let mut vm=service_vm();
    let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    let compiled=service.compile_source_by_name(
        "<?php namespace StatementService; $r=stream_service_late(3);function stream_service_late($x){return $x+4;}return $r;",
        "php","/private/tmp/vybe-statement-service-entry.php").unwrap();
    numeric_service_value(service.run_compiled(compiled).unwrap(),7);
}
#[test]
fn registered_php_public_service_dynamic_include_preserves_caller_scope() {
    struct Scratch(std::path::PathBuf);
    impl Drop for Scratch {fn drop(&mut self){let _=std::fs::remove_dir_all(&self.0);}}
    let dir=Scratch(std::env::temp_dir().join(format!("vybe-statement-include-{}-{}",
        std::process::id(),std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap().as_nanos())));
    std::fs::create_dir(&dir.0).unwrap();
    let child=dir.0.join("child.php");std::fs::write(&child,"<?php $scope+=2;return 3;").unwrap();
    for (kind,expected) in [("include",14),("include_once",10),("require",14),("require_once",10)] {
        // First execution must observe caller scope and return 9. A second
        // ordinary include executes again; *_once returns true without effects.
        let source=format!("<?php $scope=4;$answer={kind} '{}';if($scope+$answer!==9)return -1;$again={kind} '{}';return $scope+$answer+$again;",
            child.display(),child.display());
        let mut vm=service_vm();
        let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
        let compiled=service.compile_source_by_name(source,"php",dir.0.join(format!("{kind}.php"))).unwrap();
        numeric_service_value(service.run_compiled(compiled).unwrap(),expected);
    }
}

#[test]
fn registered_php_public_service_eval_preserves_scope_and_late_declaration() {
    let mut vm=service_vm();
    let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    let source=r#"<?php $scope=4;$answer=eval('$scope+=2;function statement_stream_eval_late($x){return $x+1;}return statement_stream_eval_late(2);');return $scope+$answer;"#;
    let compiled=service.compile_source_by_name(source,"php","/private/tmp/vybe-statement-eval-entry.php").unwrap();
    numeric_service_value(service.run_compiled(compiled).unwrap(),9);
}
#[test]
fn malformed_public_submission_preserves_prior_vm_and_publishes_no_partial_code() {
    let mut vm=service_vm();
    let language=vybe_compiler::languages::find_by_name("php").unwrap();
    let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    numeric_service_value(service.submit_source(
        "<?php $GLOBALS['statement_stream_survivor']=7;return 7;",language,
        "/private/tmp/vybe-statement-submit-first.php").unwrap(),7);
    let installed=service.vm().chunks.len();
    let failed=service.compile_submission_source(
        "<?php $GLOBALS['statement_stream_leak']=9;function broken(",language,
        "/private/tmp/vybe-statement-submit-broken.php");
    assert!(failed.is_err(),"malformed source must not become a compiled unit");
    assert_eq!(service.vm().chunks.len(),installed,"grammar failure installed provisional code");
    numeric_service_value(service.submit_source(
        "<?php if(isset($GLOBALS['statement_stream_leak']))return -1;return $GLOBALS['statement_stream_survivor'];",language,
        "/private/tmp/vybe-statement-submit-after.php").unwrap(),7);
}
