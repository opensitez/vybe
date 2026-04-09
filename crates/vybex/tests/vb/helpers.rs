use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_host::gui_state::GuiState;

/// Run VB source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_vb(src: &str) -> Vec<String> {
    let module = vybex::languages::vb::parse(src).expect("VB parse failed");

    let profile = load_vb_profile();

    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    let result = output.lock().unwrap().clone();
    result
}

/// Run VB source, return (VM, output) for post-run inspection of globals etc.
pub fn run_vb_vm(src: &str) -> (VM, Arc<Mutex<Vec<String>>>) {
    let module = vybex::languages::vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    (vm, output)
}

/// Run VB source with GUI host functions, return (VM, GuiState, output).
/// Uses register_all_with_gui which creates widgets directly (no side effects).
pub fn run_vb_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>) {
    let module = vybex::languages::vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    (vm, gui, output)
}

pub fn load_vb_profile() -> vybex::profile::LanguageProfile {
    use vybex::profile::*;
    use std::collections::HashMap;

    LanguageProfile {
        function_return: ReturnStyle::ResultSlot,
        self_keyword: "me".into(),
        base_keyword: Some("mybase".into()),
        constructor_name: "New".into(),
        separated_methods: false,
        implicit_self_fields: true,
        explicit_self_param: false,
        enum_as_ordinals: false,
        case_sensitive: false,
        string_indexing: StringIndexing::OneBased,
        array_upper_bound_inclusive: true,
        parens_for_index: true,
        entry_point: Some("main".into()),
        hoist_var: false,
        dynamic_add: false,
        commonjs_require: false,
        partial_classes: true,
        byref_boxing: true,
        with_block: true,
        event_side_effects: false,
        new_with_initializer: true,
        new_from_initializer: true,
        linq_queries: true,
        builtins: {
            let mut b = HashMap::new();
            b.insert("console.writeline".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
            b.insert("console.write".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
            b.insert("msgbox".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 1, max_args: 3 });
            b.insert("cstr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
            b.insert("cint".into(), BuiltinDef { emit: BuiltinEmit::Opcode("i32_from_f64".into()), min_args: 1, max_args: 1 });
            b.insert("cdbl".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_from_i32".into()), min_args: 1, max_args: 1 });
            b.insert("cbool".into(), BuiltinDef { emit: BuiltinEmit::Opcode("dyn_to_bool".into()), min_args: 1, max_args: 1 });
            b.insert("val".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "val".into()), min_args: 1, max_args: 1 });
            b.insert("len".into(), BuiltinDef { emit: BuiltinEmit::StrLength, min_args: 1, max_args: 1 });
            b.insert("abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_abs".into()), min_args: 1, max_args: 1 });
            b.insert("sqr".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_sqrt".into()), min_args: 1, max_args: 1 });
            b.insert("round".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_nearest".into()), min_args: 1, max_args: 1 });
            b.insert("fix".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_floor".into()), min_args: 1, max_args: 1 });
            b.insert("int".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_floor".into()), min_args: 1, max_args: 1 });
            b.insert("math.abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_abs".into()), min_args: 1, max_args: 1 });
            b.insert("math.floor".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_floor".into()), min_args: 1, max_args: 1 });
            b.insert("math.sqrt".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_sqrt".into()), min_args: 1, max_args: 1 });
            b.insert("math.round".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_nearest".into()), min_args: 1, max_args: 1 });
            b.insert("math.min".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_min".into()), min_args: 2, max_args: 2 });
            b.insert("math.max".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_max".into()), min_args: 2, max_args: 2 });
            b.insert("ucase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_to_upper".into()), min_args: 1, max_args: 1 });
            b.insert("lcase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_to_lower".into()), min_args: 1, max_args: 1 });
            b.insert("trim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_trim".into()), min_args: 1, max_args: 1 });
            b.insert("left".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("left".into()), min_args: 2, max_args: 2 });
            b.insert("right".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "right".into()), min_args: 2, max_args: 2 });
            b.insert("mid".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("mid".into()), min_args: 2, max_args: 3 });
            b.insert("instr".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("instr".into()), min_args: 2, max_args: 3 });
            b.insert("replace".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("replace".into()), min_args: 3, max_args: 3 });
            b.insert("split".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("split".into()), min_args: 1, max_args: 2 });
            b.insert("join".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("join".into()), min_args: 1, max_args: 2 });
            b.insert("chr".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_from_char_code".into()), min_args: 1, max_args: 1 });
            b.insert("asc".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("asc".into()), min_args: 1, max_args: 1 });
            b.insert("ubound".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("ubound".into()), min_args: 1, max_args: 1 });
            b.insert("isnothing".into(), BuiltinDef { emit: BuiltinEmit::Opcode("ref_is_null".into()), min_args: 1, max_args: 1 });
            b.insert("typename".into(), BuiltinDef { emit: BuiltinEmit::Opcode("ref_typeof".into()), min_args: 1, max_args: 1 });
            b.insert("isnumeric".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "isNumeric".into()), min_args: 1, max_args: 1 });
            b.insert("space".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("space".into()), min_args: 1, max_args: 1 });
            b.insert("ltrim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_trim_start".into()), min_args: 1, max_args: 1 });
            b.insert("rtrim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_trim_end".into()), min_args: 1, max_args: 1 });
            b.insert("lbound".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("lbound".into()), min_args: 0, max_args: 1 });
            b.insert("strreverse".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_reverse".into()), min_args: 1, max_args: 1 });
            b.insert("string".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("string_repeat".into()), min_args: 2, max_args: 2 });
            b
        },
        intrinsics: {
            let mut i = HashMap::new();
            i.insert("ubound".into(), "array_len, const_i32(1), i32_sub".into());
            i.insert("asc".into(), "const_i32(0), str_char_code_at".into());
            i.insert("left".into(), "const_i32(0), swap, to_int, str_substring".into());
            i.insert("mid".into(), "mid_1based".into());
            i.insert("instr".into(), "str_index_of, const_i32(1), i32_add".into());
            i.insert("replace".into(), "str_replace".into());
            i.insert("split".into(), "str_split".into());
            i.insert("join".into(), "array_join".into());
            i.insert("space".into(), "const_str( ), swap, to_int, str_repeat".into());
            i.insert("lbound".into(), "const_i32(0)".into());
            i.insert("string_repeat".into(), "swap, str_repeat".into());
            i
        },
        namespaces: NamespaceConfig {
            use_dotnet: true,
            extra_imports: vec!["microsoft.visualbasic".into()],
            ..Default::default()
        },
        known_types: HashMap::new(),
        value_methods: HashMap::new(),
        module_aliases: HashMap::new(),
        namespace_constants: HashMap::new(),
        array_methods: HashMap::new(),
    }
}
