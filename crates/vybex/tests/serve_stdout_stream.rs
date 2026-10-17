//! Served stdout uses the canonical consumer transport, including writes after transfer.
use std::sync::{Arc, Mutex};
use vybe_runtime::VM;
use vybex::server::script as adapter;
fn request() -> (
    Arc<vybe_platform_node::http::context::RequestContext>,
    std::sync::mpsc::Receiver<vybe_platform_node::http::context::ResponseMessage>,
) {
    use vybe_platform_node::http::context::*;
    let (tx, rx) = std::sync::mpsc::channel();
    (
        Arc::new(RequestContext {
            method: "GET".into(),
            uri: "/probe.php".into(),
            path: "/probe.php".into(),
            query: String::new(),
            scheme: "http".into(),
            host: "localhost".into(),
            headers: vec![],
            env: Default::default(),
            remote_addr: "127.0.0.1".into(),
            remote_port: 1234,
            body: Mutex::new(RequestBodyReader::empty()),
            response: Mutex::new(ResponseState::new(Some(tx))),
        }),
        rx,
    )
}
fn compile(writes: &[&str]) -> Vec<vybe_runtime::Chunk> {
    let mut chunk = vybe_runtime::Chunk::new("serve-stdout-production-transport");
    let slot = chunk.alloc_scratch(1);
    for text in writes {
        chunk.emit_string_const(text, 1);
        chunk.emit_op_u16(vybe_runtime::Op::LOCAL_SET, slot, 1);
        vybe_compiler::primitives::io::emit_write_stdout_slot(&mut chunk, slot, 1);
    }
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 1);
    chunk.emit_op(vybe_runtime::Op::RETURN, 1);
    vec![chunk]
}
fn vm() -> VM {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms(
        &mut vm,
        &vybe_runtime::capabilities::Capabilities::all(),
    );
    adapter::register_response_stdout(&mut vm);
    vm
}
fn drain(
    rx: &std::sync::mpsc::Receiver<vybe_platform_node::http::context::ResponseMessage>,
) -> Vec<u8> {
    use vybe_platform_node::http::context::ResponseMessage;
    let mut headers = 0;
    let mut body = Vec::new();
    for msg in rx.try_iter() {
        match msg {
            ResponseMessage::Headers { status, .. } => {
                assert_eq!(status, 200);
                headers += 1;
            }
            ResponseMessage::Data(bytes) => {
                assert_eq!(headers, 1);
                body.extend(bytes);
            }
        }
    }
    assert_eq!(headers, 1);
    body
}
#[test]
fn producer_after_sink_reaches_response_and_reused_vm_keeps_tenants_separate() {
    let mut vm = vm();
    for expected in ["first-tenant", "second-tenant"] {
        let (ctx, rx) = request();
        let _guard = vybe_platform_node::http::context::install_context(ctx.clone());
        vm.run(compile(&[expected])).unwrap();
        assert_eq!(drain(&rx), expected.as_bytes());
        assert!(ctx.response.lock().unwrap().headers_sent);
    }
}
#[test]
fn multiple_writes_and_unicode_keep_exact_bytes() {
    let (ctx, rx) = request();
    let _guard = vybe_platform_node::http::context::install_context(ctx.clone());
    let mut vm = vm();
    vm.run(compile(&["日本語", " / ", "second"])).unwrap();
    assert_eq!(drain(&rx), "日本語 / second".as_bytes());
}
#[test]
fn large_output_is_not_lost() {
    let expected = "abcdef".repeat(20_000);
    let (ctx, rx) = request();
    let _guard = vybe_platform_node::http::context::install_context(ctx.clone());
    let mut vm = vm();
    vm.run(compile(&[&expected])).unwrap();
    assert_eq!(drain(&rx), expected.as_bytes());
}
#[test]
fn context_retention_survives_nested_tenant_switch() {
    let (first, _) = request();
    let (second, _) = request();
    let _a = vybe_platform_node::http::context::install_context(first.clone());
    let retained = vybe_platform_node::http::context::current_context().unwrap();
    {
        let _b = vybe_platform_node::http::context::install_context(second.clone());
        assert!(Arc::ptr_eq(&retained, &first));
        assert!(Arc::ptr_eq(
            &vybe_platform_node::http::context::current_context().unwrap(),
            &second
        ));
    }
    assert!(Arc::ptr_eq(
        &vybe_platform_node::http::context::current_context().unwrap(),
        &first
    ));
}

#[test]
fn response_bytes_arrive_before_script_completion() {
    let (ctx, rx) = request();
    let _guard = vybe_platform_node::http::context::install_context(ctx.clone());
    let mut vm = vm();
    let observed = Arc::new(std::sync::atomic::AtomicBool::new(false));
    let seen = observed.clone();
    let receiver = Mutex::new(rx);
    vm.register_host_fn(
        "serve:test",
        "checkpoint",
        Box::new(move |_, _| {
            assert_eq!(drain(&receiver.lock().unwrap()), b"early-output");
            assert!(!ctx.response.lock().unwrap().ended);
            seen.store(true, std::sync::atomic::Ordering::Relaxed);
            vybe_runtime::Value::Null
        }),
    );
    let mut chunks = compile(&["early-output"]);
    let chunk = &mut chunks[0];
    // Insert the observation before RETURN, after the ordinary output transport.
    chunk.code.truncate(chunk.code.len() - 4);
    chunk.lines.truncate(chunk.lines.len() - 4);
    let checkpoint = chunk.add_import("serve:test", "checkpoint");
    chunk.emit_call(checkpoint, 0, 1);
    chunk.emit_op(vybe_runtime::Op::DROP, 1);
    chunk.emit_op(vybe_runtime::Op::RETURN, 1);
    vm.run(chunks).unwrap();
    assert!(observed.load(std::sync::atomic::Ordering::Relaxed));
}
