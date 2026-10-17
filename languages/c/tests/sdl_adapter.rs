//! Headless SDL sidecar. No Doom translation units or desktop runner are used.
//! Windowed operation bodies run synchronously in a test-only AST; untouched
//! production emission, bounded worker completion and DOM keyboard delivery
//! are separate gates.
//! `--nocapture inventory` prints each registered API's source-call fixtures
//! and compatibility-stub status. Call presence is not a correctness claim.
use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use vybe_ast::{ExprKind, StmtKind};
use vybe_runtime::{Chunk, Op, VM, Value};

static SERIAL: Mutex<()> = Mutex::new(());

const THREADED_FIXTURES: &[&str] = &[
    "sdl_adapter_thread_completion",
    "sdl_adapter_thread_keyboard",
];

const FAMILIES: &[(&str, &[&str])] = &[
    (
        "surface",
        &[
            "sdl_adapter_borrowed_surface",
            "sdl_rgba_surface_blit",
            "sdl_blit_pitch_masks",
            "sdl_indexed_to_rgba_full_blit",
            "libc_bulk_sdl_byte_equivalence",
            "sdl_linear_palette",
            "sdl_palette_lookup_pointer",
            "sdl_pixel_word_view",
        ],
    ),
    (
        "window_geometry",
        &[
            "sdl_adapter_window_geometry",
            "sdl_adapter_display_geometry",
            "sdl_renderer_info_output",
        ],
    ),
    (
        "renderer_texture",
        &[
            "sdl_adapter_texture_target",
            "sdl_lock_texture_member_outputs",
            "sdl_render_target_scaled_copy",
            "sdl_present_backbuffer",
            "sdl_paletted_texture_blit",
            "sdl_render_copy_paletted_void_branch",
        ],
    ),
    (
        "input",
        &[
            "sdl_adapter_mouse_queue",
            "sdl_peep_event_queue",
            "sdl_keyboard_event_roundtrip",
            "sdl_keyboard_translation_table",
            "sdl_named_scancode_keycodes",
            "sdl_modifier_comparison",
            "sdl_poll_wait_queue",
            "endoom_wait_event_contract",
        ],
    ),
    (
        "lifecycle_clock",
        &[
            "sdl_video_init_return_status",
            "sdl_external_declaration_init",
            "sdl_boolean_constants",
            "sdl_tick_comparison",
            "sdl_timer_return_status",
            "sdl_adapter_performance_clock",
            "sdl_delay_void_branch",
            "sdl_delay_nested_stack",
            "sdl_get_pref_path_fallback",
            "sdl_pref_path_config_join",
        ],
    ),
    (
        "canvas_dialog",
        &[
            "sdl_adapter_canvas_commands",
            "sdl_adapter_renderer_fill_rect",
            "sdl_adapter_palette_canvas",
            "sdl_error_message",
        ],
    ),
    (
        "compatibility",
        &[
            "sdl_adapter_compatibility",
            "sdl_adapter_audio_thread_stubs",
            "sdl_adapter_typed_string_read",
            "sdl_adapter_boolean_zero",
        ],
    ),
];

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup")
        .join(format!("{name}.c"))
}

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

fn validate(chunks: &[Chunk]) -> Result<(), String> {
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .map(|_| ())
        .map_err(|error| error.to_string())
}

fn synchronous_operations(
    mut compiled: vybe_compiler::dynamic::DynamicCompilation,
    vm: &VM,
    path: &Path,
) -> Result<(vybe_compiler::dynamic::DynamicCompilation, bool), String> {
    if compiled.app_shell.is_none() {
        return Ok((compiled, false));
    }
    let bundle = vybe_compiler::projects::load(path)?;
    let mut module = bundle.prepared_module()?;
    let mut guests = 0;
    for statement in &mut module.body {
        statement.walk_exprs_mut(&mut |expression| {
            let ExprKind::Call { callee, args, .. } = &expression.kind else {
                return;
            };
            if !matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_thread_spawn") {
                return;
            }
            assert_eq!(args.len(), 2, "SDL worker ABI changed");
            let ExprKind::FunctionExpr(guest) = &args[1].value.kind else {
                panic!("SDL worker no longer carries its lowered guest body");
            };
            assert!(
                matches!(&guest.kind, StmtKind::FunctionDecl { name, .. } if name == "__sdl_guest")
            );
            guests += 1;
            use vybe_platform_libc::emitter::build::*;
            let invocation = call_expr(args[1].value.clone(), vec![args[0].value.clone()]);
            *expression = expr(ExprKind::Sequence(vec![
                assign_expr(ident("__sdl_sidecar_guest_started"), int_lit(1)),
                invocation,
            ]));
        });
    }
    assert_eq!(
        guests, 1,
        "windowed operation fixture must execute one guest"
    );
    module.body.insert(
        0,
        vybe_platform_libc::emitter::build::typed_var_decl_stmt(
            "__sdl_sidecar_guest_started",
            "int",
            vybe_platform_libc::emitter::build::int_lit(0),
        ),
    );
    let synchronous = bundle.compile_prepared_module(&module, &vm.modules)?;
    validate(&synchronous.chunks)?;
    compiled.chunks = synchronous.chunks;
    compiled.host_imports = synchronous.host_imports;
    compiled.required_libraries = synchronous.required_libraries;
    Ok((compiled, true))
}

fn batch(family: &str) {
    let _guard = SERIAL.lock().unwrap_or_else(|error| error.into_inner());
    let names = FAMILIES.iter().find(|(name, _)| *name == family).unwrap().1;
    let mut failures = Vec::new();
    for name in names {
        let mut vm = runtime();
        let compiled = match vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&fixture(name))
        {
            Ok(compiled) => compiled,
            Err(error) => {
                failures.push(format!("{name}: compile: {error}"));
                continue;
            }
        };
        let calls = Arc::new(Mutex::new(Vec::<(String, Vec<Value>)>::new()));
        let mut recorders = VM::new();
        // Observe only standard WHATWG boundaries; never create native dialogs.
        if matches!(*name, "sdl_adapter_canvas_commands" | "sdl_adapter_renderer_fill_rect") {
            for operation in [
                "setFillStyle",
                "fillRect",
                "setStrokeStyle",
                "beginPath",
                "moveTo",
                "lineTo",
                "stroke",
                "fillText",
            ] {
                let calls = calls.clone();
                recorders.register_host_fn(
                    "web:canvas",
                    operation,
                    Box::new(move |_, args| {
                        calls
                            .lock()
                            .unwrap()
                            .push((operation.into(), args[1..].to_vec()));
                        Value::Undefined
                    }),
                );
            }
        }
        if *name == "sdl_error_message" {
            let calls = calls.clone();
            recorders.register_host_fn(
                "web:window",
                "alert",
                Box::new(move |_, args| {
                    calls.lock().unwrap().push(("alert".into(), args.to_vec()));
                    Value::Undefined
                }),
            );
        }
        if *name == "sdl_adapter_palette_canvas" {
            let calls = calls.clone();
            recorders.register_host_fn(
                "web:canvas",
                "putImageData",
                Box::new(move |_, args| {
                    calls
                        .lock()
                        .unwrap()
                        .push(("putImageData".into(), args[1..].to_vec()));
                    Value::Undefined
                }),
            );
        }
        // Preserve the existing host indexes used by typed exports and aliases.
        vm.overlay_host_fns_from(&recorders);
        match validate(&compiled.chunks) {
            Ok(()) => eprintln!("{name}: production metadata-free WASM PASS"),
            Err(error) => failures.push(format!("{name}: WASM: {error}")),
        }
        // Production chunks above stay untouched. Operation tests run the exact
        // lowered guest synchronously; these are not scheduler verification.
        let (compiled, synchronous) = match synchronous_operations(compiled, &vm, &fixture(name)) {
            Ok(compiled) => compiled,
            Err(error) => {
                failures.push(format!("{name}: synchronous harness: {error}"));
                continue;
            }
        };
        // Keep execution and validation gates independent, including failures.
        let source = std::fs::read_to_string(fixture(name)).unwrap();
        let expected_exit = source
            .lines()
            .find_map(|line| {
                line.strip_prefix("// vybe-test-exit:").map(|exit| {
                    exit.trim()
                        .parse::<i32>()
                        .expect("invalid fixture exit status")
                })
            })
            .unwrap_or(0);
        match vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm).run_compiled(compiled) {
            Ok(_) if vm.pending_exit_code == expected_exit => {
                eprintln!("{name}: operation PASS (synchronous guest={synchronous})");
            }
            Ok(_) => failures.push(format!(
                "{name}: exit {}, expected {expected_exit}",
                vm.pending_exit_code
            )),
            Err(error) => failures.push(format!("{name}: source: {error}")),
        }
        let calls = calls.lock().unwrap();
        if synchronous && vm.global("__sdl_sidecar_guest_started").map(Value::as_i32) != Some(1) {
            failures.push(format!("{name}: guest entry was not executed"));
        }
        if *name == "sdl_adapter_palette_canvas" {
            if calls.len() != 2 {
                failures.push(format!("{name}: expected two image uploads"));
            }
            for (_, args) in calls.iter() {
                let bytes = args.first().and_then(image_bytes);
                if bytes.as_deref() != Some(&[129, 17, 255, 255, 3, 193, 41, 255][..])
                    || args.len() != 3
                    || args[1].as_i32() != 0
                    || args[2].as_i32() != 0
                {
                    failures.push(format!("{name}: palette image mismatch: {bytes:?}"));
                }
            }
        }
        if *name == "sdl_adapter_canvas_commands" {
            let expected = [
                ("setFillStyle", vec![17, 34, 51, 255]),
                ("fillRect", vec![3, 5, 7, 9]),
                ("setStrokeStyle", vec![68, 85, 102, 255]),
                ("beginPath", vec![]),
                ("moveTo", vec![2, 4]),
                ("lineTo", vec![6, 8]),
                ("stroke", vec![]),
                ("setFillStyle", vec![119, 136, 153, 255]),
            ];
            if calls.len() != 9 {
                failures.push(format!(
                    "{name}: expected 9 canvas calls, got {}",
                    calls.len()
                ));
            }
            for ((operation, args), (wanted, values)) in calls.iter().zip(expected) {
                if operation != wanted
                    || args.iter().map(Value::as_i32).collect::<Vec<_>>() != values
                {
                    failures.push(format!(
                        "{name}: canvas call mismatch {operation}: {args:?}"
                    ));
                }
            }
            if let Some((operation, args)) = calls.get(8) {
                if operation != "fillText"
                    || args.len() != 3
                    || !matches!(&args[0], Value::String(text) if text.as_ref() == "headless")
                    || args[1].as_i32() != 10
                    || args[2].as_i32() != 12
                {
                    failures.push(format!("{name}: text call mismatch: {args:?}"));
                }
            }
        }
        if *name == "sdl_adapter_renderer_fill_rect" {
            let expected = [
                ("setFillStyle", vec![17, 34, 51, 255]),
                ("fillRect", vec![3, 5, 7, 9]),
                ("setFillStyle", vec![68, 85, 102, 255]),
                ("fillRect", vec![0, 0, 32, 24]),
            ];
            if calls.len() != expected.len() {
                failures.push(format!("{name}: expected 4 canvas calls, got {}", calls.len()));
            }
            for ((operation, args), (wanted, values)) in calls.iter().zip(expected) {
                if operation != wanted
                    || args.iter().map(Value::as_i32).collect::<Vec<_>>() != values
                {
                    failures.push(format!("{name}: canvas call mismatch {operation}: {args:?}"));
                }
            }
        }
        if *name == "sdl_error_message"
            && (calls.len() != 3
                || calls
                    .iter()
                    .any(|(_, args)| args.len() != 1 || !matches!(&args[0], Value::String(_))))
        {
            failures.push(format!(
                "{name}: expected three standard alert string calls"
            ));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

fn image_bytes(image: &Value) -> Option<Vec<u8>> {
    use vybe_runtime::value::ObjectKind;
    let Value::Object(image) = image else {
        return None;
    };
    let data = image.lock().unwrap().properties.get("data")?.clone();
    let Value::Object(data) = data else {
        return None;
    };
    let data = data.lock().unwrap();
    let ObjectKind::TypedArray(array) = &data.kind else {
        return None;
    };
    let buffer = array.buffer.lock().unwrap();
    buffer
        .get(array.byte_offset..array.byte_offset.checked_add(array.length)?)
        .map(|bytes| bytes.to_vec())
}

#[test]
fn surfaces() {
    batch("surface");
}
#[test]
fn window_and_display_geometry() {
    batch("window_geometry");
}
#[test]
fn renderer_and_textures() {
    batch("renderer_texture");
}
#[test]
fn input_queue_and_translation() {
    batch("input");
}
#[test]
fn lifecycle_and_clocks() {
    batch("lifecycle_clock");
}
#[test]
fn canvas_commands_and_dialog_strings() {
    batch("canvas_dialog");
}
#[test]
fn compatibility_source_contracts() {
    batch("compatibility");
}

#[test]
fn production_thread_reaches_terminal_guest_sentinel() {
    let _guard = SERIAL.lock().unwrap_or_else(|error| error.into_inner());
    let mut vm = runtime();
    let path = fixture("sdl_adapter_thread_completion");
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path)
        .unwrap();
    validate(&compiled.chunks).expect("production thread module must validate");
    assert!(
        compiled.chunks.iter().any(|chunk| chunk
            .imports
            .iter()
            .any(|import| import.module == "wasi:threads" && import.name == "thread-spawn")),
        "integration fixture must retain the production scheduler path"
    );
    let (sender, receiver) = std::sync::mpsc::channel();
    let mut recorder = VM::new();
    recorder.register_host_fn(
        "web:window",
        "alert",
        Box::new(move |_, args| {
            sender
                .send(args.to_vec())
                .expect("thread sentinel receiver alive");
            Value::I32(0)
        }),
    );
    vm.overlay_host_fns_from(&recorder);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .expect("production parent entry");
    let args = receiver
        .recv_timeout(std::time::Duration::from_secs(5))
        .expect("guest did not reach its terminal sentinel within five seconds");
    assert_eq!(args.len(), 1);
    let Value::String(message) = &args[0] else {
        panic!("dialog sentinel must be text");
    };
    assert!(
        message.contains("guest-completed"),
        "unexpected sentinel {message}"
    );
    // This establishes terminal body progress, not native UI or worker-error
    // propagation. Synchronous operation gates independently check exit status.
}

#[test]
fn production_thread_receives_dom_keyboard_events() {
    use engine::{DomOp, DomValue, UiEventFields};
    use std::sync::{Condvar, mpsc};
    use std::time::Duration;
    use vybe_platform_web::{engine, html};

    // Release on assertion failure as well as success, without leaving a
    // worker blocked inside the recorder boundary.
    struct Release(Arc<(Mutex<bool>, Condvar)>);
    impl Release {
        fn release(&self) {
            *self.0.0.lock().unwrap() = true;
            self.0.1.notify_all();
        }
    }
    impl Drop for Release {
        fn drop(&mut self) {
            self.release();
        }
    }
    fn field(event: &Value, name: &str) -> Option<Value> {
        let Value::Object(event) = event else {
            return None;
        };
        event.lock().unwrap().properties.get(name).cloned()
    }
    fn check_fields(
        event: &Value,
        expected: &UiEventFields,
        phase: &str,
        failures: &mut Vec<String>,
    ) {
        for (name, wanted) in [
            ("type", Value::String(expected.kind.clone().into())),
            ("key", Value::String(expected.key.clone().into())),
            ("code", Value::String(expected.code.clone().into())),
            ("keyCode", Value::I32(expected.key_code)),
            ("ctrlKey", Value::Bool(expected.ctrl_key)),
        ] {
            let actual = field(event, name);
            if actual != Some(wanted.clone()) {
                failures.push(format!(
                    "{phase}.{name}: expected {wanted:?}, got {actual:?}"
                ));
            }
        }
    }

    let _guard = SERIAL.lock().unwrap_or_else(|error| error.into_inner());
    vybe_platform_web::engine_select::choose(vybe_platform_web::engine_select::Engine::WebCore);
    let mut vm = runtime();
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&fixture("sdl_adapter_thread_keyboard"))
        .unwrap();
    validate(&compiled.chunks).expect("production keyboard module must validate");
    let input = compiled
        .chunks
        .iter()
        .find(|chunk| chunk.name == "__libc_sdl_input")
        .expect("production SDL callback emitted");
    eprintln!(
        "SDL input callback: arity={}, takes_receiver={}",
        input.arity, input.takes_receiver
    );
    assert!(
        compiled.chunks.iter().any(|chunk| chunk
            .imports
            .iter()
            .any(|import| import.module == "wasi:threads" && import.name == "thread-spawn")),
        "keyboard integration must retain the production scheduler"
    );
    let release = Release(Arc::new((Mutex::new(false), Condvar::new())));
    let barrier = release.0.clone();
    let (sender, receiver) = mpsc::channel();
    let mut recorder = VM::new();
    recorder.register_host_fn(
        "web:window",
        "alert",
        Box::new(move |_, args| {
            let message = match args {
                [Value::String(message)] => message.to_string(),
                _ => format!("invalid dialog arguments: {args:?}"),
            };
            let ready = message.contains("keyboard-ready");
            let _ = sender.send(message);
            if ready {
                let open = barrier.0.lock().unwrap();
                let (open, timeout) = barrier
                    .1
                    .wait_timeout_while(open, Duration::from_secs(5), |open| !*open)
                    .unwrap();
                if timeout.timed_out() && !*open {
                    let _ = sender.send("keyboard-barrier-timeout".into());
                }
            }
            Value::I32(0)
        }),
    );
    vm.overlay_host_fns_from(&recorder);
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .expect("production keyboard parent entry");
    let ready = receiver
        .recv_timeout(Duration::from_secs(5))
        .expect("guest keyboard readiness");
    assert!(
        ready.contains("keyboard-ready"),
        "unexpected readiness: {ready}"
    );

    let document = html::active_document();
    let DomValue::Node(canvas) = engine::apply(document, DomOp::QuerySelector("canvas".into()))
    else {
        panic!("guest window must create a real canvas");
    };
    engine::apply(
        document,
        DomOp::SetAttribute(canvas, "tabindex".into(), "0".into()),
    );
    engine::apply(document, DomOp::Focus(canvas));
    for kind in ["keydown", "keyup"] {
        assert_eq!(
            html::listeners_for(document, engine::DOCUMENT, kind).len(),
            1,
            "production SDL listener for {kind}"
        );
    }
    let mut expected = Vec::new();
    for (key, code, key_code) in [
        ("ArrowUp", "ArrowUp", 38),
        ("Control", "ControlRight", 17),
        ("a", "KeyA", 65),
    ] {
        for kind in ["keydown", "keyup"] {
            let event = UiEventFields {
                kind: kind.into(),
                key: key.into(),
                code: code.into(),
                key_code,
                ctrl_key: code == "ControlRight" && kind == "keydown",
                ..UiEventFields::default()
            };
            engine::apply(document, DomOp::DispatchKeyboard(event.clone()));
            expected.push(event);
        }
    }
    let mut failures = Vec::new();
    let pending = html::pending_dispatches(document);
    if pending.len() != expected.len() {
        failures.push(format!(
            "DOM callback count: expected {}, got {}",
            expected.len(),
            pending.len()
        ));
    }
    for (index, ((callback, event), expected)) in pending.iter().zip(&expected).enumerate() {
        check_fields(event, expected, &format!("DOM[{index}]"), &mut failures);
        for (name, wanted) in [("target", canvas), ("currentTarget", engine::DOCUMENT)] {
            if field(event, name).map(|value| value.as_f64()) != Some(wanted as f64) {
                failures.push(format!("DOM[{index}].{name}: {:?}", field(event, name)));
            }
        }
        // Same single-event invocation as the production GUI loop. No manual
        // receiver argument and no SDL_PushEvent or synthetic queue insertion.
        if let Err(error) = vm.invoke(callback, std::slice::from_ref(event)) {
            failures.push(format!("VM callback[{index}]: {error}"));
        }
    }
    let queue = match vm.global(vybe_platform_libc::emitter::sdl_app::INPUT_QUEUE) {
        Some(Value::Object(queue)) => {
            let queue = queue.lock().unwrap();
            match &queue.kind {
                vybe_runtime::value::ObjectKind::Array(events) => events.clone(),
                _ => {
                    failures.push("SDL queue is not an array".into());
                    Vec::new()
                }
            }
        }
        other => {
            failures.push(format!("SDL queue missing: {other:?}"));
            Vec::new()
        }
    };
    if queue.len() != expected.len() {
        failures.push(format!(
            "SDL queue length: expected {}, got {}",
            expected.len(),
            queue.len()
        ));
    }
    for (index, (event, expected)) in queue.iter().zip(&expected).enumerate() {
        check_fields(
            event,
            expected,
            &format!("SDL queue[{index}]"),
            &mut failures,
        );
    }
    release.release();
    let terminal = receiver.recv_timeout(Duration::from_secs(5));
    if !matches!(&terminal, Ok(message) if message.contains("keyboard-completed")) {
        failures.push(format!(
            "guest SDL_PollEvent terminal sentinel: {terminal:?}"
        ));
    }
    eprintln!(
        "threaded keyboard: production WASM validation=PASS; DOM callbacks={}; SDL queued={}; terminal={terminal:?}; failures={failures:?}",
        pending.len(),
        queue.len()
    );
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

// These describe CURRENT adapter contracts, not native SDL semantics. In
// particular, byte swaps, physical input and device APIs are compatibility stubs.
fn stub_result(name: &str) -> Option<Value> {
    if matches!(name, "SDL_GetPrefPath") {
        return Some(Value::String("./vybe/".into()));
    }
    if matches!(
        name,
        "SDL_GetError"
            | "SDL_GetKeyName"
            | "Mix_GetError"
            | "SDLNet_GetError"
            | "SDL_JoystickName"
            | "SDL_JoystickNameForIndex"
            | "SDL_GameControllerName"
            | "SDL_GameControllerMappingForGUID"
    ) {
        return Some(Value::String("".into()));
    }
    if matches!(
        name,
        "SDL_JoystickOpen"
            | "SDL_GameControllerOpen"
            | "SDL_CreateThread"
            | "Mix_LoadMUS"
            | "Mix_LoadMUS_RW"
            | "SDLNet_UDP_Open"
            | "SDLNet_AllocPacket"
    ) {
        return Some(Value::Null);
    }
    if matches!(
        name,
        "SDL_SetHint"
            | "SDL_SetHintWithPriority"
            | "SDL_GetAppState"
            | "SDL_GetNumVideoDisplays"
            | "SDL_GetWindowID"
    ) {
        return Some(Value::I32(1));
    }
    if name.starts_with("Mix_")
        || name.starts_with("SDLNet_")
        || name.starts_with("SDL_Joystick")
        || name.starts_with("SDL_GameController")
        || matches!(
            name,
            "SDL_NumJoysticks"
                | "SDL_IsGameController"
                | "SDL_LockAudio"
                | "SDL_UnlockAudio"
                | "SDL_PauseAudio"
                | "SDL_BuildAudioCVT"
                | "SDL_ConvertAudio"
                | "SDL_MixAudioFormat"
                | "SDL_CreateMutex"
                | "SDL_DestroyMutex"
                | "SDL_LockMutex"
                | "SDL_UnlockMutex"
                | "SDL_CreateCond"
                | "SDL_DestroyCond"
                | "SDL_CondWait"
                | "SDL_CondSignal"
                | "SDL_WaitThread"
                | "SDL_UpdateWindowSurfaceRects"
                | "SDL_QuitSubSystem"
                | "SDL_free"
                | "SDL_FreeSurface"
                | "SDL_DestroyRenderer"
                | "SDL_DestroyTexture"
                | "SDL_GetWindowDisplayIndex"
                | "SDL_GetWindowFlags"
                | "SDL_SetWindowTitle"
                | "SDL_SetWindowFullscreen"
                | "SDL_SetWindowIcon"
                | "SDL_SetWindowMinimumSize"
                | "SDL_WarpMouseInWindow"
                | "SDL_SetRelativeMouseMode"
                | "SDL_GetRelativeMouseState"
                | "SDL_StartTextInput"
                | "SDL_StopTextInput"
                | "SDL_SetTextInputRect"
                | "SDL_IsTextInputActive"
                | "SDL_RenderSetLogicalSize"
                | "SDL_RenderSetIntegerScale"
                | "SDL_UnlockTexture"
                | "SDL_LockSurface"
                | "SDL_UnlockSurface"
                | "SDL_PumpEvents"
        )
    {
        return Some(Value::I32(0));
    }
    None
}

#[test]
fn every_registered_compatibility_stub_has_an_independent_emission_case() {
    let _guard = SERIAL.lock().unwrap_or_else(|error| error.into_inner());
    let mut api_count = 0;
    let mut module_count = 0;
    for &(name, _) in vybe_platform_libc::emitter::tree_register::SDL_COMMON_EMITS {
        let swap = matches!(
            name,
            "SDL_SwapLE16" | "SDL_SwapLE32" | "SDL_SwapBE16" | "SDL_SwapBE32"
        );
        let Some(expected) = stub_result(name).or_else(|| swap.then_some(Value::I32(129))) else {
            continue;
        };
        api_count += 1;
        for (prefix, argc) in [("sdl", 0), ("sdl", 1), ("sdl", 3), ("libc.sdl", 1)] {
            let mut chunks = vec![Chunk::new("<script>")];
            for argument in 0..argc {
                chunks[0].emit_i32_const(129 + i32::from(argument), 0);
            }
            assert!(
                vybe_platform_libc::emitter::sdl::emit_sdl(
                    &format!("{prefix}.{name}"),
                    &mut chunks,
                    0,
                    argc,
                    0
                ),
                "{name}"
            );
            chunks[0].emit_op(Op::RETURN, 0);
            validate(&chunks).unwrap_or_else(|error| panic!("{name}: {error}"));
            let result = runtime()
                .run(chunks)
                .unwrap_or_else(|error| panic!("{name}: {error}"));
            let expected = if swap && argc == 0 {
                Value::I32(0)
            } else {
                expected.clone()
            };
            assert_eq!(result, expected, "{prefix}.{name}/{argc}");
            module_count += 1;
        }
    }
    eprintln!(
        "Compatibility emission: {api_count} named APIs, {module_count} independently validated/executed modules"
    );
}

// Explicit emitter controls bypass frontend pointer normalization, so a CELL
// holding NULL/zero reaches the SDL boundary rather than disappearing upstream.
fn poll_wait_test_pointer(chunk: &mut Chunk, mode: usize) {
    use vybe_compiler::primitives::class_slots::{self, ObjSource, ResolvedSlot, ValueSource};
    match mode {
        0 => chunk.emit_i32_const(0, 0),
        1 => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        2 | 5 => {
            let new = chunk.add_import("ecma:object", "new");
            chunk.emit_call(new, 0, 0);
        }
        _ => {
            let new = chunk.add_import("ecma:object", "new");
            chunk.emit_call(new, 0, 0);
            let cell = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, cell, 0);
            chunk.emit_string_const(if mode >= 6 { "cstruct" } else { "cell" }, 0);
            class_slots::emit_class_set(chunk, ObjSource::Local(cell),
                &ResolvedSlot::Key("__ref_kind".into()), ValueSource::Stack, 0);
            if mode == 3 || mode == 6 {
                chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
            } else {
                chunk.emit_i32_const(0, 0);
            }
            class_slots::emit_class_set(chunk, ObjSource::Local(cell),
                &ResolvedSlot::Key(if mode >= 6 { "__base" } else { "__value" }.into()),
                ValueSource::Stack, 0);
            chunk.emit_op_u16(Op::LOCAL_GET, cell, 0);
        }
    }
}

#[test]
fn wait_event_repolls_after_standard_clock_wait_and_null_keeps_event() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    use vybe_compiler::primitives::{collections, globals};
    use vybe_runtime::value::{Object, ObjectKind};

    let _guard = SERIAL.lock().unwrap_or_else(|error| error.into_inner());
    for mode in 0..8 {
        let mut chunks = vec![Chunk::new("<wait-event-clock-control>")];
        let queue_name = vybe_platform_libc::emitter::sdl_app::INPUT_QUEUE;
        collections::emit_array_new(&mut chunks, 0, 0, 0);
        globals::emit_write(&mut chunks[0], queue_name, 0);
        poll_wait_test_pointer(&mut chunks[0], mode);
        let destination = chunks[0].alloc_scratch(1);
        chunks[0].emit_op_u16(Op::LOCAL_SET, destination, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, destination, 0);
        globals::emit_write(&mut chunks[0], "__sdl_wait_test_destination", 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, destination, 0);
        vybe_platform_libc::emitter::sdl::emit_sdl_wait_event(&mut chunks, 0, 1, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        // Raw emitter chunks need the same global index finalization as source compilation.
        globals::finalize_global_table(&mut chunks);
        validate(&chunks).expect("wait loop must emit metadata-free valid WASM");

        let waits = Arc::new(AtomicUsize::new(0));
        let observed = waits.clone();
        let mut vm = runtime();
        let mut clock = VM::new();
        // Test-only replacement of the existing standard clock import. No new
        // production event API: make one private queue entry arrive on sleep.
        clock.register_host_fn(
            "wasi:clocks/monotonic-clock", "wait-for",
            Box::new(move |context, args| {
                assert_eq!(observed.fetch_add(1, Ordering::SeqCst), 0,
                    "a second wait means the first wakeup was lost");
                assert_eq!(args.last().unwrap().as_f64(), 1_000_000.0);
                let mut event = Object::new();
                event.set("__sdl_pushed_type".into(), if mode % 2 == 0 {
                    Value::I32(0x100)
                } else { Value::F64(256.0) });
                let Value::Object(queue) = context.get_global(queue_name) else {
                    panic!("SDL private queue must exist before waiting");
                };
                queue.lock().unwrap().set("0".into(),
                    Value::Object(vybe_runtime::heap::alloc(event)));
                let (future, id) = context.create_future();
                context.resolve_future(id, Value::Null);
                future
            }),
        );
        vm.overlay_host_fns_from(&clock);
        assert_eq!(vm.run(chunks).unwrap().as_i32(), 1);
        assert_eq!(waits.load(Ordering::SeqCst), 1);
        let Value::Object(queue) = vm.global(queue_name).unwrap() else {
            panic!("SDL private queue must remain installed");
        };
        let queue = queue.lock().unwrap();
        let ObjectKind::Array(entries) = &queue.kind else { panic!("queue is an array") };
        assert_eq!(entries.len(), if mode == 2 || mode == 5 { 0 } else { 1 });
        if mode == 2 || mode == 5 {
            let Value::Object(destination) = vm.global("__sdl_wait_test_destination").unwrap()
                else { panic!("non-NULL output record must survive the call") };
            assert_eq!(destination.lock().unwrap().get("type").as_i32(), 0x100);
        }
    }
}

#[test]
fn inventory() {
    let mut calls: BTreeMap<String, BTreeSet<&str>> = BTreeMap::new();
    for name in FAMILIES
        .iter()
        .flat_map(|(_, names)| names.iter().copied())
        .chain(THREADED_FIXTURES.iter().copied())
    {
        let text = std::fs::read_to_string(fixture(name)).unwrap();
        // Source-call inventory only. Assertions and execution are separate gates.
        for word in text.split(|c: char| !(c.is_ascii_alphanumeric() || c == '_')) {
            if word.starts_with("SDL_") {
                calls.entry(word.into()).or_default().insert(name);
            }
        }
    }
    let mut names: BTreeSet<&str> = vybe_platform_libc::emitter::tree_register::SDL_COMMON_EMITS
        .iter()
        .map(|(name, _)| *name)
        .collect();
    // This constructor is provided by C AST lowering, not SDL_COMMON_EMITS.
    names.insert("SDL_CreateRGBSurface");
    let mut missing = Vec::new();
    let mut unmapped = Vec::new();
    for name in names {
        let source = calls
            .get(name)
            .map(|names| names.iter().copied().collect::<Vec<_>>().join(", "));
        let emission = if name.starts_with("SDL_Swap") {
            "C endian conversion; independent emitter identity-compatibility test"
        } else if stub_result(name).is_some() {
            "independent compatibility-stub emission test"
        } else {
            "adapter / AST constructor"
        };
        eprintln!(
            "{name} | {emission} | {}",
            source.as_deref().unwrap_or("SOURCE GAP")
        );
        if source.is_none() {
            missing.push(name);
            if stub_result(name).is_none() && !name.starts_with("SDL_Swap") {
                unmapped.push(name);
            }
        }
    }
    eprintln!(
        "Explicit source gaps (not full API coverage): {}",
        missing.join(", ")
    );
    eprintln!(
        "Wildcard Mix_/SDLNet_/joystick/controller dispatch is not an exhaustive API contract."
    );
    eprintln!("SDL_HasIntersection/SDL_IntersectRect are not supported by this inventory.");
    assert!(
        unmapped.is_empty(),
        "Nonstub APIs without source cases: {}",
        unmapped.join(", ")
    );
}
