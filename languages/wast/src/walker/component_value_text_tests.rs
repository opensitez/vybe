use super::parse;
use vybe_ast::canon::{
    ComponentBinary, ComponentBinaryDef as Def, ComponentInterfaceDecl as Decl,
    ComponentInterfaceRef as Ref, ComponentItemKind as Kind, ValSpec as Val,
};

fn binary(source: &str) -> ComponentBinary {
    let module = parse(source).unwrap_or_else(|e| panic!("{source}: {e}"));
    assert!(module.canon.component_source);
    assert!(
        module.body.is_empty(),
        "binary AST must not become WAST runtime shims"
    );
    (*module.canon.binary.expect("typed component AST")).clone()
}

fn bytes(source: &str) -> Vec<u8> {
    let binary = binary(source);
    let bytes = vybe_platform_wasm::write_component_binary(&binary).unwrap();
    let read = vybe_platform_wasm::read_component_binary(&bytes).unwrap();
    assert_eq!(
        bytes,
        vybe_platform_wasm::write_component_binary(&read).unwrap()
    );
    bytes
}

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|c| u8::from_str_radix(std::str::from_utf8(c).unwrap(), 16).unwrap())
        .collect()
}

#[test]
fn primary_bound_exact_bytes() {
    assert_eq!(
        bytes("(component (import \"input\" (value $v u32)) (export \"output\" (value $v)))"),
        hex("0061736d0d0001000a0b010005696e7075740201790b0c0100066f7574707574020000")
    );
}

#[test]
fn start_source_names_exact_bytes() {
    assert_eq!(
        bytes(
            "(component (type (func (param \"x\" u32) (result u32))) (import \"run\" (func (type 0))) (import \"input\" (value u32)) (start 0 (value 0) (result (value))) (export \"out\" (value 1)))"
        ),
        hex(
            "0061736d0d000100070801400101787900790a1202000372756e01000005696e7075740201790904000100010b090100036f7574020100"
        )
    );
}

#[test]
fn nested_source_names_exact_bytes() {
    assert_eq!(
        bytes("(component (component (import \"x\" (value u32)) (export \"x\" (value 0))))"),
        hex("0061736d0d000100041a0061736d0d0001000a07010001780201790b0701000178020000")
    );
}

#[test]
fn literal_exact_bytes() {
    assert_eq!(
        bytes("(component (value $n u32 42) (export \"out\" (value $n)))"),
        hex("0061736d0d0001000c040179012a0b090100036f7574020000")
    );
}

#[test]
fn named_type_without_value_binder_is_not_eaten_by_optional_id() {
    let b = binary("(component (type $t u32) (value $t 5) (export \"out\" (value 0)))");
    assert!(
        matches!(&b.definitions[1], Def::ValueLiteral { ty: Val::Ref(0), payload } if payload == &[5])
    );
    let b =
        binary("(component (type $t u32) (import \"v\" (value $t)) (export \"out\" (value 0)))");
    assert!(matches!(
        &b.definitions[1],
        Def::NamedImport {
            ty: Ref::Value(Val::Ref(0)),
            ..
        }
    ));
}

#[test]
fn start_result_binders_and_rebound_value_space() {
    let b = binary(
        "(component (import \"run\" (func $run (param \"x\" u32) (result u32)))
        (import \"input\" (value $input u32))
        (start $run (value $input) (result (value $result)))
        (export $out \"out\" (value $result))
        (type $signature (instance (export \"marker\" (value u32)))))",
    );
    assert!(
        b.definitions
            .iter()
            .any(|d| matches!(d, Def::Start { function: 0, args, results: 1 } if args == &[0]))
    );
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::NamedExport {
            kind: Kind::Value,
            index: 1,
            ..
        }
    )));
    assert!(parse("(component (import \"v\" (value $v u32)) (export $e \"v\" (value $v)) (export \"again\" (value $e)))").is_err());
}

#[test]
fn independent_equal_typed_values_stay_distinct() {
    let b = binary(
        "(component (import \"a\" (value $a u32)) (import \"b\" (value $b u32))
        (instance (export \"a\" (value $a)) (export \"b\" (value $b))))",
    );
    assert!(
        matches!(b.definitions.last(), Some(Def::NamedInstanceFromExports(items)) if items[0].index == 0 && items[1].index == 1)
    );
}

#[test]
fn instance_value_alias_and_path_sugar() {
    let b = binary(
        "(component (import \"v\" (value $v u32)) (instance $i (export \"v\" (value $v)))
        (alias export $i \"v\" (value $a)) (export \"a\" (value $a))
        (export \"b\" (value $i \"v\")))",
    );
    let aliases: Vec<_> = b
        .definitions
        .iter()
        .filter(|d| {
            matches!(
                d,
                Def::AliasInstanceExport {
                    kind: Kind::Value,
                    ..
                }
            )
        })
        .collect();
    assert_eq!(aliases.len(), 2);
    assert!(matches!(
        b.definitions.last(),
        Some(Def::NamedExport { index: 3, .. })
    ));
}

#[test]
fn equality_bound_is_not_a_type_bound() {
    let b = binary(
        "(component (type (instance (export \"a\" (value $a u32)) (export \"b\" (value (eq $a))))))",
    );
    let Def::InterfaceType { declarations, .. } = &b.definitions[0] else {
        panic!()
    };
    assert!(matches!(
        &declarations[1],
        Decl::NamedExport {
            ty: Ref::ValueEq(0),
            ..
        }
    ));
    assert!(bytes("(component (type (instance (export \"a\" (value $a u32)) (export \"b\" (value (eq $a))))))")
        .windows(3).any(|w| w == [2, 0, 0]));
}

#[test]
fn value_export_ascription_and_attributes_are_retained() {
    let b = binary(
        "(component (import \"input\" (external-id \"in\") (value $v u32))
        (export \"output\" (external-id \"out\") (value $v) (value u32)))",
    );
    assert!(
        matches!(&b.definitions[0], Def::NamedImport { name, .. } if name.attributes.len() == 1)
    );
    assert!(
        matches!(b.definitions.last(), Some(Def::NamedExport { name, ty: Some(Ref::Value(Val::Prim(t))), .. }) if name.attributes.len() == 1 && t == "u32")
    );
}

#[test]
fn interface_scopes_capture_outer_types_without_value_leakage() {
    let b = binary(
        "(component $root (type $t u32)
        (type (instance (export \"v\" (value $t))))
        (component (value $v $t 7) (export \"v\" (value $v))))",
    );
    assert!(
        matches!(&b.definitions[1], Def::InterfaceType { declarations, .. }
        if matches!(&declarations[0], Decl::AliasOuterType { count: 1, index: 0 }))
    );
    let Def::NestedComponent(child) = &b.definitions[2] else {
        panic!()
    };
    assert!(matches!(&child.definitions[0], Def::AliasOuter(a) if a.count == 1 && a.index == 0));
    assert!(parse("(component (value $outer u8 1) (component (export \"v\" (value $outer))) (export \"v\" (value $outer)))").is_err());
}

#[test]
fn multiple_starts_are_primary_not_old_tool_policy() {
    let b = binary("(component (import \"run\" (func $run)) (start $run) (start $run))");
    assert_eq!(
        b.definitions
            .iter()
            .filter(|d| matches!(d, Def::Start { .. }))
            .count(),
        2
    );
}

fn payload(ty: &str, value: &str) -> Vec<u8> {
    let b = binary(&format!(
        "(component (type $t {ty}) (export $public \"t\" (type $t))
        (value $v $public {value}) (export \"v\" (value $v) (value $public)))"
    ));
    b.definitions
        .iter()
        .find_map(|d| {
            if let Def::ValueLiteral { payload, .. } = d {
                Some(payload.clone())
            } else {
                None
            }
        })
        .unwrap()
}

#[test]
fn literal_integer_payloads_and_natural_ranges() {
    for (ty, value, expected) in [
        ("s8", "-128", vec![128]),
        ("s16", "-129", vec![255, 126]),
        ("u16", "65535", vec![255, 255, 3]),
        (
            "u64",
            "18446744073709551615",
            vec![255, 255, 255, 255, 255, 255, 255, 255, 255, 1],
        ),
    ] {
        assert_eq!(payload(ty, value), expected);
    }
    for (ty, value) in [
        ("s8", "128"),
        ("s16", "65535"),
        ("s64", "9223372036854775808"),
        ("u8", "-1"),
        ("u16", "65536"),
    ] {
        assert!(
            parse(&format!(
                "(component (value $v {ty} {value}) (export \"v\" (value $v)))"
            ))
            .is_err(),
            "{ty}: {value}"
        );
    }
}

#[test]
fn literal_floats_are_canonical_and_width_rounded() {
    assert_eq!(payload("f32", "nan"), [0, 0, 192, 127]);
    assert_eq!(payload("f64", "nan"), [0, 0, 0, 0, 0, 0, 248, 127]);
    assert_eq!(payload("f32", "0x1.8p+1"), 3_f32.to_le_bytes());
    for value in ["-nan", "nan:0x1", "1e1000"] {
        assert!(
            parse(&format!(
                "(component (value $v f32 {value}) (export \"v\" (value $v)))"
            ))
            .is_err()
        );
    }
}

#[test]
fn literal_utf8_and_composite_payloads() {
    assert_eq!(payload("char", "'a'"), [97]);
    assert_eq!(payload("string", "\"hi\""), [2, 104, 105]);
    assert_eq!(payload("(list u8)", "(list 1 2)"), [2, 1, 2]);
    assert_eq!(payload("(list u8 2)", "(list 1 2)"), [1, 2]);
    assert_eq!(payload("(tuple u8 u16)", "(tuple 1 128)"), [1, 128, 1]);
    assert_eq!(payload("(record (field \"a\" u8))", "(record 7)"), [7]);
    assert_eq!(
        payload(
            "(variant (case \"a\") (case \"b\" u8))",
            "(variant \"b\" 7)"
        ),
        [1, 7]
    );
    assert_eq!(payload("(option u8)", "none"), [0]);
    assert_eq!(payload("(result u8)", "(ok 7)"), [0, 7]);
    assert_eq!(
        payload("(flags \"a\" \"b\" \"c\")", "(flags \"c\" \"a\")"),
        [5]
    );
    assert_eq!(payload("(enum \"a\" \"b\")", "(enum \"b\")"), [1]);
    assert_eq!(
        payload("(map u8 u16)", "(map (entry 1 128) (entry 2 3))"),
        [2, 1, 128, 1, 2, 3]
    );
}

#[test]
fn literal_binary_form_is_native_validated() {
    assert_eq!(payload("u32", "(binary \"\\2a\")"), [42]);
    for (ty, value) in [
        ("bool", "(binary \"\\02\")"),
        ("u32", "(binary \"\\80\")"),
        ("u8", "(binary \"\\00\\00\")"),
        ("(option (stream u8))", "none"),
        ("char", "'\\u{d800}'"),
        ("char", "'\\u{110000}'"),
        ("(list error-context)", "(list)"),
        ("(list u8 2)", "(list 1)"),
        ("(flags \"a\")", "(flags \"a\" \"a\")"),
    ] {
        assert!(
            parse(&format!(
                "(component (value $v {ty} {value}) (export \"v\" (value $v) (value {ty})))"
            ))
            .is_err(),
            "{ty}: {value}"
        );
    }
}

#[test]
fn comments_do_not_change_literal_shape() {
    assert_eq!(
        payload("( (; type ;) tuple u8)", "( (; value ;) tuple 9)"),
        [9]
    );
}

#[test]
fn non_value_and_core_module_behavior_is_preserved() {
    let m = parse("(module (func (export \"f\") (result i32) i32.const 7))").unwrap();
    assert!(!m.canon.component_source);
    assert!(!m.body.is_empty());
    let m = parse("(component (type $t u32) (export \"t\" (type $t)))").unwrap();
    assert!(m.canon.binary.is_some());
    let b = binary(
        "(component (core module (func (export \"f\") (result i32) i32.const 7))
        (value $v u8 1) (export \"v\" (value $v)))",
    );
    assert!(matches!(&b.definitions[0], Def::CoreModule(bytes) if bytes.len() > 8));
}

#[test]
fn invalid_assertions_use_native_value_validation() {
    let m = parse("(assert_invalid (component (import \"unused\" (value u32))) \"component value is unused\")").unwrap();
    assert!(
        !m.body
            .iter()
            .any(|s| matches!(s.kind, vybe_ast::StmtKind::Throw { .. }))
    );
}

#[test]
fn executable_value_scripts_are_not_silently_lowered() {
    assert!(parse("(component (value $v u8 1) (export \"v\" (value $v))) (invoke \"f\")").is_err());
}

#[test]
fn mixed_core_module_import_and_inline_type_are_native() {
    let b = binary(
        "(component
        (core type $Schema (module (export \"run\" (func (result i32)))))
        (import \"module\" (core module $Imported (type $Schema)))
        (core instance $i (instantiate $Imported))
        (canon lift (core func $i \"run\") (func $run (result u32)))
        (value u8 1) (export \"v\" (value 0)) (export \"run\" (func $run)))",
    );
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::NamedImport {
            ty: Ref::Module(0),
            ..
        }
    )));
    assert!(
        b.definitions
            .iter()
            .any(|d| matches!(d, Def::CoreInstance { module: 0, .. }))
    );
    bytes("(component (import \"empty\" (core module $m)) (value u8 1) (export \"v\" (value 0)))");
}

#[test]
fn mixed_core_recursive_groups_keep_member_slots_and_outer_identity() {
    let b = binary(
        "(component $Root
        (core rec (type $A (func)) (type $B (func (param i32))))
        (type $B u8)
        (component $Child
            (alias outer $Root $B (core type $Copied))
            (value $v $B 7) (export \"v\" (value $v))))",
    );
    assert!(matches!(&b.definitions[0],Def::CoreType(ty) if ty.type_count()==Some(2)));
    let child = b
        .definitions
        .iter()
        .find_map(|d| {
            if let Def::NestedComponent(c) = d {
                Some(c)
            } else {
                None
            }
        })
        .unwrap();
    let aliases: Vec<_> = child
        .definitions
        .iter()
        .filter_map(|d| d.normalized_outer_alias())
        .collect();
    assert!(aliases.iter().any(
        |a| a.kind == vybe_ast::canon::ComponentOuterAliasKind::CoreType
            && a.count == 1
            && a.index == 1
    ));
    assert!(
        aliases
            .iter()
            .any(|a| a.kind == vybe_ast::canon::ComponentOuterAliasKind::Type
                && a.count == 1
                && a.index == 0)
    );
}

#[test]
fn mixed_resource_destructor_alias_precedes_later_core_functions() {
    let b = binary("(component
        (core module $M (func (export \"drop\") (param i32)) (func (export \"get\") (result i32) i32.const 7))
        (core instance $i (instantiate $M))
        (type $R (resource (rep i32) (dtor (core func $i \"drop\"))))
        (alias core export $i \"get\" (core func $get))
        (canon lift (core func $get) (func $read (result u32)))
        (start $read (result (value $v))) (export \"v\" (value $v)))");
    assert!(
        b.definitions
            .iter()
            .any(|d| matches!(d, Def::ResourceType { dtor: Some(0), .. }))
    );
    assert!(
        b.definitions
            .iter()
            .any(|d| matches!(d, Def::CanonLift { core_func: 1, .. }))
    );
}

#[test]
fn mixed_inline_component_argument_has_its_own_instance_slot() {
    let b = binary(
        "(component
        (component $Child
            (type $I (instance (export \"v\" (value u8))))
            (import \"i\" (instance (type $I))))
        (value u8 1)
        (instance (instantiate $Child (with \"i\" (instance (export \"v\" (value 0)))))))",
    );
    assert!(b.definitions.iter().any(|d|matches!(d,Def::ComponentInstance{args,..} if args.len()==1&&args[0].kind==Kind::Instance&&args[0].index==0)));
}

#[test]
fn interface_only_value_types_preserve_executable_walks() {
    let source = "(component
        (core module $M (func (export \"run\") (result i32) i32.const 42))
        (core instance $i (instantiate $M))
        (func $f (result u32) (canon lift (core func $i \"run\")))
        (export \"run\" (func $f))
        (type (instance (export \"marker\" (value u32)))))
        (assert_return (invoke \"run\") (u32.const 42))";
    let module = parse(source).unwrap();
    assert!(!module.body.is_empty());
    assert_eq!(module.canon.defs.len(), 1);
    assert!(
        !module
            .body
            .iter()
            .any(|s| matches!(s.kind, vybe_ast::StmtKind::Throw { .. }))
    );
}

#[test]
fn primary_equality_ascriptions_use_native_exact_evidence() {
    assert_eq!(
        bytes("(component (value $v u32 42) (export \"v\" (value $v) (value (eq $v))))"),
        hex("0061736d0d0001000c040179012a0b0a01000176020001020000")
    );
    assert!(
        parse(
            "(component (value $a u32 42) (value $b u32 43)
        (export \"a\" (value $a) (value (eq $b))) (export \"b\" (value $b)))"
        )
        .is_err()
    );
}

#[test]
fn related_instance_value_fields_preserve_source_value_indices() {
    let b = binary(
        "(component
        (type $I (instance (export \"a\" (value $a u32)) (export \"b\" (value (eq $a)))))
        (import \"i\" (instance $i (type $I)))
        (alias export $i \"a\" (value $a)) (alias export $i \"b\" (value $b))
        (export \"a\" (value $a) (value (eq $b))) (export \"b\" (value $b)))",
    );
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::NamedExport {
            kind: Kind::Value,
            index: 0,
            ty: Some(Ref::ValueEq(1)),
            ..
        }
    )));
}

#[test]
fn binary_literal_uses_native_instance_type_alias_resolution() {
    let b = binary(
        "(component
        (type $I (instance (type $t u32) (export \"t\" (type (eq $t)))))
        (import \"i\" (instance $i (type $I))) (alias export $i \"t\" (type $t))
        (value $v $t (binary \"\\2a\")) (export \"v\" (value $v) (value $t)))",
    );
    assert!(
        b.definitions
            .iter()
            .any(|d| matches!(d,Def::ValueLiteral{ty:Val::Ref(1),payload} if payload==&[42]))
    );
}

#[test]
fn unaliased_instance_value_fields_keep_the_native_binary() {
    let b = binary(
        "(component (type $I (instance (export \"v\" (value u32))))
        (import \"i\" (instance (type $I))))",
    );
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::NamedImport {
            ty: Ref::Instance(0),
            ..
        }
    )));
}

fn aliased_literal_pair(schema: &str, text: &str, payload: &[u8]) -> ComponentBinary {
    aliased_literal_pair_with_declarations("", schema, text, payload)
}

fn aliased_literal_pair_with_declarations(
    declarations: &str,
    schema: &str,
    text: &str,
    payload: &[u8],
) -> ComponentBinary {
    let prefix=format!("(component (type $I (instance {declarations} (type $t {schema}) (export \"payload\" (type (eq $t)))))
        (import \"i\" (instance $i (type $I))) (alias export $i \"payload\" (type $selected))");
    let binary = literal_source_pair(&prefix, "$selected", text, payload);
    assert!(
        binary
            .definitions
            .iter()
            .any(|d| matches!(d,Def::ValueLiteral{ty:Val::Ref(1),payload:p} if p==payload))
    );
    binary
}

fn literal_source_pair(prefix: &str, ty: &str, text: &str, payload: &[u8]) -> ComponentBinary {
    let source = |value: &str| {
        format!("{prefix} (value $v {ty} {value}) (export \"out\" (value $v) (value {ty})))")
    };
    let encoded: String = payload.iter().map(|b| format!("\\{b:02x}")).collect();
    let raw = format!("(binary \"{encoded}\")");
    let a = bytes(&source(text));
    let b = bytes(&source(&raw));
    assert_eq!(
        a, b,
        "same source declarations must emit identical primary bytes"
    );
    if let Some(root) = std::env::var_os("VALUE_TEXT_PROOF_DIR") {
        use std::hash::{Hash, Hasher};
        let mut hash = std::collections::hash_map::DefaultHasher::new();
        source(text).hash(&mut hash);
        let root = std::path::PathBuf::from(root);
        std::fs::create_dir_all(&root).unwrap();
        let id = format!("{:016x}", hash.finish());
        for (suffix, data) in [
            ("text.wat", source(text).into_bytes()),
            ("binary.wat", source(&raw).into_bytes()),
            ("text.wasm", a),
            ("binary.wasm", b),
            ("payload.bin", payload.to_vec()),
        ] {
            std::fs::write(root.join(format!("{id}.{suffix}")), data).unwrap();
        }
    }
    binary(&source(text))
}

#[test]
fn instance_exported_primitive_shapes_have_exact_text_binary_bytes() {
    for (ty, text, payload) in [
        ("u32", "300", vec![172, 2]),
        ("s16", "-129", vec![255, 126]),
        ("string", "\"hi\"", vec![2, 104, 105]),
        ("char", "'a'", vec![97]),
        ("bool", "true", vec![1]),
        ("u8", "255", vec![255]),
        ("u16", "65535", vec![255, 255, 3]),
        (
            "u64",
            "18446744073709551615",
            vec![255, 255, 255, 255, 255, 255, 255, 255, 255, 1],
        ),
        ("s8", "-128", vec![128]),
        ("s32", "-129", vec![255, 126]),
        ("s64", "-129", vec![255, 126]),
        ("f32", "nan", vec![0, 0, 192, 127]),
        ("f64", "nan", vec![0, 0, 0, 0, 0, 0, 248, 127]),
        ("char", "'\\u{1f600}'", vec![240, 159, 152, 128]),
    ] {
        aliased_literal_pair(ty, text, &payload);
    }
}

#[test]
fn instance_exported_compound_shapes_have_exact_text_binary_bytes() {
    for (ty, text, payload) in [
        (
            "(record (field \"left\" u8) (field \"right\" (list u16)))",
            "(record 7 (list 128 2))",
            vec![7, 2, 128, 1, 2],
        ),
        (
            "(variant (case \"none\") (case \"some\" (tuple u8 bool)))",
            "(variant \"some\" (tuple 7 true))",
            vec![1, 7, 1],
        ),
        (
            "(option (result u8 (error u16)))",
            "(some (error 128))",
            vec![1, 1, 128, 1],
        ),
        ("(list u8 3)", "(list 1 2 3)", vec![1, 2, 3]),
        ("(list u8)", "(list 1 2 3)", vec![3, 1, 2, 3]),
        ("(enum \"one\" \"two\")", "(enum \"two\")", vec![1]),
        ("(flags \"one\" \"two\")", "(flags \"two\")", vec![2]),
    ] {
        aliased_literal_pair(ty, text, &payload);
    }
    aliased_literal_pair_with_declarations(
        "(type $flags (flags \"one\" \"two\")) (export \"flags\" (type (eq $flags)))",
        "(map string $flags)",
        "(map (entry \"a\" (flags \"two\")))",
        &[1, 1, 97, 2],
    );
}

#[test]
fn aliased_constructor_checks_are_not_weakened_by_type_projection() {
    for (schema, value) in [
        ("u8", "256"),
        ("char", "'\\u{d800}'"),
        ("char", "'\\u{110000}'"),
        ("(tuple u8 bool)", "(tuple 1)"),
        ("(variant (case \"some\" u8))", "(variant \"Some\" 1)"),
        ("(list u8 2)", "(list 1)"),
        ("(option (stream u8))", "none"),
    ] {
        let source = format!(
            "(component (type $I (instance (type $t {schema}) (export \"payload\" (type (eq $t)))))
            (import \"i\" (instance $i (type $I))) (alias export $i \"payload\" (type $selected))
            (value $v $selected {value}) (export \"out\" (value $v) (value $selected)))"
        );
        assert!(parse(&source).is_err(), "{schema} {value}");
    }
    for value in [
        "(map (entry \"a\" (flags \"two\")))",
        "(binary \"\\01\\01a\\02\")",
    ] {
        let source=format!("(component
            (type $I (instance (type $t (map string (flags \"one\" \"two\"))) (export \"payload\" (type (eq $t)))))
            (import \"i\" (instance $i (type $I))) (alias export $i \"payload\" (type $selected))
            (value $v $selected {value}) (export \"out\" (value $v) (value $selected)))");
        let error = parse(&source)
            .expect_err("hidden anonymous flags must not become public through shape projection");
        assert!(error.contains("externally visible import name"), "{error}");
    }
}

#[test]
fn nested_instance_alias_paths_project_the_exported_type_declaration() {
    let prefix = "(component
        (type $Outer (instance
            (type $Inner (instance (type $t (tuple u8 u16)) (export \"payload\" (type (eq $t)))))
            (export \"child\" (instance (type $Inner)))))
        (import \"outer\" (instance $outer (type $Outer)))
        (alias export $outer \"child\" (instance $child))
        (alias export $child \"payload\" (type $selected))";
    let b = literal_source_pair(prefix, "$selected", "(tuple 7 300)", &[7, 172, 2]);
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::ValueLiteral {
            ty: Val::Ref(1),
            ..
        }
    )));
}

#[test]
fn instantiated_component_type_exports_support_text_and_binary_literals() {
    let prefix = "(component
        (component $Factory (type $t (record (field \"left\" u16))) (export \"payload\" (type $t)))
        (instance $i (instantiate $Factory)) (alias export $i \"payload\" (type $selected))";
    let b = literal_source_pair(prefix, "$selected", "(record 300)", &[172, 2]);
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::ValueLiteral {
            ty: Val::Ref(0),
            ..
        }
    )));
}

#[test]
fn inline_instance_and_type_export_rebinding_keep_actual_indices() {
    let prefix = "(component (type $t (tuple u8 u16)) (export $public \"tuple\" (type $t))
        (instance $i (export \"payload\" (type $public)))
        (alias export $i \"payload\" (type $first))
        (export $again \"tuple-alias\" (type $first))
        (alias export $i \"payload\" (type $selected))";
    let b = literal_source_pair(prefix, "$selected", "(tuple 7 300)", &[7, 172, 2]);
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::ValueLiteral {
            ty: Val::Ref(4),
            ..
        }
    )));
    assert_eq!(
        b.definitions
            .iter()
            .filter(|d| matches!(
                d,
                Def::AliasInstanceExport {
                    kind: Kind::Type,
                    ..
                }
            ))
            .count(),
        2
    );
}

#[test]
fn identical_export_labels_in_separate_interfaces_do_not_select_the_wrong_shape() {
    let prefix="(component
        (type $A (instance (type $t u8) (export \"payload\" (type (eq $t)))))
        (type $B (instance (type $t u16) (export \"payload\" (type (eq $t)))))
        (import \"a\" (instance $a (type $A))) (import \"b\" (instance $b (type $B)))
        (alias export $a \"payload\" (type $narrow)) (alias export $b \"payload\" (type $selected))";
    let b = literal_source_pair(prefix, "$selected", "300", &[172, 2]);
    assert!(b.definitions.iter().any(|d| matches!(
        d,
        Def::ValueLiteral {
            ty: Val::Ref(3),
            ..
        }
    )));
    assert!(
        parse(&format!(
            "{prefix} (value $v $narrow 300) (export \"out\" (value $v)))"
        ))
        .is_err()
    );
}

#[test]
fn shared_lexical_start_result_origins_do_not_prove_runtime_value_equality() {
    let prefix = "(component (import \"run\" (func $run (result u32)))
        (component $Factory (import \"run\" (func $run (result u32)))
            (start $run (result (value $result))) (export \"result\" (value $result)))
        (instance $a (instantiate $Factory (with \"run\" (func $run))))";
    bytes(&format!("{prefix} (alias export $a \"result\" (value $first)) (alias export $a \"result\" (value $second))
        (export \"first\" (value $first) (value (eq $second))) (export \"second\" (value $second)))"));
    assert!(parse(&format!("{prefix}
        (instance $b (instantiate $Factory (with \"run\" (func $run))))
        (alias export $a \"result\" (value $first)) (alias export $b \"result\" (value $second))
        (export \"first\" (value $first) (value (eq $second))) (export \"second\" (value $second)))")).is_err());
}

#[test]
fn exported_instance_ascription_projects_only_the_exposed_view() {
    let prefix = "(component
        (type $Actual (instance (type $a u8) (export \"payload\" (type (eq $a)))
            (type $b u16) (export \"hidden\" (type (eq $b)))))
        (type $View (instance (type $a u8) (export \"payload\" (type (eq $a)))))
        (import \"actual\" (instance $actual (type $Actual)))
        (export $view \"view\" (instance $actual) (instance (type $View)))";
    literal_source_pair(
        &format!("{prefix} (alias export $view \"payload\" (type $selected))"),
        "$selected",
        "7",
        &[7],
    );
    assert!(
        parse(&format!(
            "{prefix} (alias export $view \"hidden\" (type $selected))
        (value $v $selected 300) (export \"out\" (value $v) (value $selected)))"
        ))
        .is_err()
    );
}

#[test]
fn outer_capture_of_an_imported_type_alias_retains_its_shape() {
    let prefix = "(component
        (type $I (instance (type $t (tuple u8 u16)) (export \"payload\" (type (eq $t)))))
        (import \"i\" (instance $i (type $I))) (alias export $i \"payload\" (type $captured))";
    let make = |value: &str| {
        format!(
            "{prefix} (component (value $v $captured {value})
        (export \"out\" (value $v) (value $captured))))"
        )
    };
    assert_eq!(
        bytes(&make("(tuple 7 300)")),
        bytes(&make("(binary \"\\07\\ac\\02\")"))
    );
}
