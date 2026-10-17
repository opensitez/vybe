use vybe_ast::{ExprKind, StmtKind};

#[test]
fn environment_seed_uses_static_integer_guards() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/environment_seed.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let helper = chunks.iter().find(|chunk| chunk.name == "__c_env_init").unwrap();
    let dump = vybe_runtime::debug::disassemble(helper);
    assert!(!dump.contains("__vybe_slot_"), "{dump}");
    assert!(!dump.contains("dyneq"), "{dump}");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn bounded_varargs_format_uses_one_static_write_helper() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/bounded_varargs_format.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let helpers = chunks.iter().filter(|chunk| chunk.name == "__c_snprintf_write")
        .collect::<Vec<_>>();
    assert_eq!(helpers.len(), 1);
    let dump = vybe_runtime::debug::disassemble(helpers[0]);
    assert!(dump.contains("i64.lt_u") || dump.contains("i64.gt_u"), "{dump}");
    assert!(!dump.contains("__vybe_slot_"), "{dump}");
    assert!(!dump.contains("dyneq"), "{dump}");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn linear_strncpy_parameters_use_bulk_copy_and_fill() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_strncpy_parameter_copy.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let dump = chunks.iter().map(vybe_runtime::debug::disassemble).collect::<Vec<_>>().join("\n");
    assert!(dump.contains("memory.copy"), "missing bulk copy");
    assert!(dump.contains("memory.fill"), "missing zero padding");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn c_option_search_uses_numeric_string_positions() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/argv_option_search.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_entry_does_not_emit_foreign_shutdown_hooks() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/sdl/paletted.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let dump = vybe_runtime::debug::disassemble(&chunks[0]);
    assert!(!dump.contains("__php_shutdown_fns"), "{dump}");
    assert!(!dump.contains("__vybe_ob_stack"), "{dump}");
}

#[test]
fn c_entry_passes_main_arguments_explicitly() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/explicit_main_arguments.c");
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_guest_keeps_the_explicit_main_call() {
    let mut module = vybe_language_c::parse(
        "#include <SDL.h>\nint main(int argc, char **argv) {\n\
         SDL_Init(SDL_INIT_VIDEO);\n\
         SDL_CreateWindow(\"test\", 0, 0, 32, 32, 0);\n\
         return argc;\n}",
    ).unwrap();
    let main = module.body.iter().find_map(|statement| match &statement.kind {
        StmtKind::FunctionDecl { name, params, .. } if name == "main" => Some(params),
        _ => None,
    }).unwrap();
    assert_eq!(main.len(), 2);
    assert!(main.iter().all(|param| param.default.is_none()));
    let mut calls = 0;
    let mut guests = 0;
    for statement in &mut module.body {
        statement.walk_exprs_mut(&mut |expression| {
            if let ExprKind::FunctionExpr(function) = &expression.kind {
                if matches!(&function.kind, StmtKind::FunctionDecl { name, .. } if name == "__sdl_guest") {
                    guests += 1;
                }
            }
            if let ExprKind::Call { callee, args, .. } = &expression.kind {
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "main") {
                    calls += 1;
                    assert_eq!(args.len(), 2);
                }
            }
        });
    }
    assert_eq!(guests, 1);
    assert_eq!(calls, 1);
}
