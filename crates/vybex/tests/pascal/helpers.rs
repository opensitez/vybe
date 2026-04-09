use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

/// Run Pascal source through vybex pipeline: pest grammar -> walker -> common AST -> compiler -> VM
pub fn run_pascal(src: &str) -> Vec<String> {
    let module = vybex::languages::pascal::parse(src).expect("Pascal parse failed");

    let profile = load_pascal_profile();

    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("Pascal compile failed");

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
    vm.run(chunks).expect("Pascal run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn load_pascal_profile() -> vybex::profile::LanguageProfile {
    use vybex::profile::*;
    use std::collections::HashMap;

    LanguageProfile {
        function_return: ReturnStyle::ResultSlot,
        self_keyword: "self".into(),
        base_keyword: Some("inherited".into()),
        constructor_name: "Create".into(),
        separated_methods: true,
        implicit_self_fields: true,
        explicit_self_param: false,
        enum_as_ordinals: true,
        case_sensitive: false,
        string_indexing: StringIndexing::OneBased,
        array_upper_bound_inclusive: false,
        parens_for_index: false,
        entry_point: Some("main".into()),
        hoist_var: false,
        dynamic_add: false,
        commonjs_require: false,
        partial_classes: false,
        byref_boxing: false,
        with_block: false,
        event_side_effects: false,
        new_with_initializer: false,
        new_from_initializer: false,
        linq_queries: false,
        builtins: {
            let mut b = HashMap::new();
            b.insert("writeln".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
            b.insert("write".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
            b.insert("length".into(), BuiltinDef { emit: BuiltinEmit::StrLength, min_args: 1, max_args: 1 });
            b.insert("inttostr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
            b.insert("floattostr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
            b.insert("strtoint".into(), BuiltinDef { emit: BuiltinEmit::Opcode("i32_from_f64".into()), min_args: 1, max_args: 1 });
            b.insert("strtofloat".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "val".into()), min_args: 1, max_args: 1 });
            b.insert("abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_abs".into()), min_args: 1, max_args: 1 });
            b.insert("sqr".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("sqr".into()), min_args: 1, max_args: 1 });
            b.insert("sqrt".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_sqrt".into()), min_args: 1, max_args: 1 });
            b.insert("round".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_nearest".into()), min_args: 1, max_args: 1 });
            b.insert("trunc".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_trunc".into()), min_args: 1, max_args: 1 });
            b.insert("floor".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_floor".into()), min_args: 1, max_args: 1 });
            b.insert("ceil".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_ceil".into()), min_args: 1, max_args: 1 });
            b.insert("min".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_min".into()), min_args: 2, max_args: 2 });
            b.insert("max".into(), BuiltinDef { emit: BuiltinEmit::Opcode("f64_max".into()), min_args: 2, max_args: 2 });
            b.insert("power".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "pow".into()), min_args: 2, max_args: 2 });
            b.insert("uppercase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_to_upper".into()), min_args: 1, max_args: 1 });
            b.insert("lowercase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_to_lower".into()), min_args: 1, max_args: 1 });
            b.insert("trim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_trim".into()), min_args: 1, max_args: 1 });
            b.insert("concat".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("concat".into()), min_args: 1, max_args: 255 });
            b.insert("copy".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("copy".into()), min_args: 3, max_args: 3 });
            b.insert("pos".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("pos".into()), min_args: 2, max_args: 2 });
            b.insert("stringreplace".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("replace".into()), min_args: 3, max_args: 3 });
            b.insert("stringofchar".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("stringofchar".into()), min_args: 2, max_args: 2 });
            b.insert("leftstr".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("left".into()), min_args: 2, max_args: 2 });
            b.insert("rightstr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "right".into()), min_args: 2, max_args: 2 });
            b.insert("chr".into(), BuiltinDef { emit: BuiltinEmit::Opcode("str_from_char_code".into()), min_args: 1, max_args: 1 });
            b.insert("ord".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("asc".into()), min_args: 1, max_args: 1 });
            b.insert("inc".into(), BuiltinDef { emit: BuiltinEmit::MutateVar("add".into()), min_args: 1, max_args: 2 });
            b.insert("dec".into(), BuiltinDef { emit: BuiltinEmit::MutateVar("sub".into()), min_args: 1, max_args: 2 });
            b.insert("succ".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("succ".into()), min_args: 1, max_args: 1 });
            b.insert("pred".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("pred".into()), min_args: 1, max_args: 1 });
            b.insert("high".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("high".into()), min_args: 1, max_args: 1 });
            b.insert("low".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("low".into()), min_args: 1, max_args: 1 });
            b.insert("assigned".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("assigned".into()), min_args: 1, max_args: 1 });
            b.insert("freeandnil".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("freeandnil".into()), min_args: 1, max_args: 1 });
            b.insert("maxint".into(), BuiltinDef { emit: BuiltinEmit::Intrinsic("maxint".into()), min_args: 0, max_args: 0 });
            b.insert("randomize".into(), BuiltinDef { emit: BuiltinEmit::Noop, min_args: 0, max_args: 0 });
            b
        },
        intrinsics: {
            let mut i = HashMap::new();
            i.insert("sqr".into(), "dup, f64_mul".into());
            i.insert("high".into(), "array_len, const_i32(1), i32_sub".into());
            i.insert("low".into(), "const_i32(0)".into());
            i.insert("asc".into(), "const_i32(0), str_char_code_at".into());
            i.insert("left".into(), "const_i32(0), swap, to_int, str_substring".into());
            i.insert("copy".into(), "copy_1based".into());
            i.insert("pos".into(), "str_index_of, const_i32(1), i32_add".into());
            i.insert("replace".into(), "str_replace".into());
            i.insert("concat".into(), "str_concat_variadic".into());
            i.insert("stringofchar".into(), "swap, to_int, str_repeat".into());
            i.insert("succ".into(), "const_i32(1), i32_add".into());
            i.insert("pred".into(), "const_i32(1), i32_sub".into());
            i.insert("assigned".into(), "ref_is_null, i32_eqz".into());
            i.insert("freeandnil".into(), "set_null".into());
            i.insert("maxint".into(), "const_i32(2147483647)".into());
            i
        },
        namespaces: NamespaceConfig::default(),
        known_types: HashMap::new(),
        value_methods: HashMap::new(),
        module_aliases: HashMap::new(),
        namespace_constants: HashMap::new(),
        array_methods: HashMap::new(),
    }
}
