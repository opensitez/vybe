//! Real current writer before vs isolated standard native packed-i16 writer.
use vybe_runtime::chunk::{CompositeKind, TypeEntry};
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, slot: u16) {
    c.emit_op_u16(Op::LOCAL_GET, slot, 0);
}
fn set(c: &mut Chunk, slot: u16) {
    c.emit_op_u16(Op::LOCAL_SET, slot, 0);
}
fn func(name: &str, argc: u8) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = argc;
    c.param_count = argc;
    c.local_count = u16::from(argc);
    c
}
fn finish(c: &mut Chunk) {
    c.emit_op(Op::RETURN, 0);
}
fn call(c: &mut Chunk, name: &str, argc: u8) {
    let i = c.add_import("wasm:js-string", name);
    c.emit_call(i, argc, 0);
}
fn module(access: bool, classes: usize) -> Vec<Chunk> {
    let mut owner = func("packed_owner", 0);
    owner.emit_i32_const(1, 0);
    finish(&mut owner);
    // Genuine declared storage metadata: kind Array and element i16, not a name hint.
    let mut types = vec![TypeEntry {
        name: "Characters".into(),
        kind: CompositeKind::Array,
        fields: vec!["i16".into()],
        ..Default::default()
    }];
    for i in 0..classes {
        types.push(TypeEntry {
            name: format!("Class{i}"),
            ..Default::default()
        });
    }
    let mut rt = func("roundtrip", 1);
    rt.local_count = 3;
    get(&mut rt, 0);
    call(&mut rt, "length", 1);
    set(&mut rt, 1);
    get(&mut rt, 1);
    rt.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 1, 0);
    set(&mut rt, 2);
    get(&mut rt, 0);
    get(&mut rt, 2);
    rt.emit_i32_const(0, 0);
    call(&mut rt, "intoCharCodeArray", 3);
    rt.emit_op(Op::DROP, 0);
    get(&mut rt, 2);
    rt.emit_i32_const(0, 0);
    get(&mut rt, 1);
    call(&mut rt, "fromCharCodeArray", 3);
    finish(&mut rt);
    let mut out = vec![owner, rt];
    if access {
        let mut c = func("packed_new", 1);
        get(&mut c, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 1, 0);
        finish(&mut c);
        out.push(c);
        let mut c = func("packed_fill_new", 2);
        get(&mut c, 0);
        get(&mut c, 1);
        c.emit_op_u16(Op::ARRAY_NEW, 1, 0);
        finish(&mut c);
        out.push(c);
        let mut c = func("packed_fixed", 0);
        for n in [0, -1, 32768, 65535, 131071] {
            c.emit_i32_const(n, 0);
        }
        c.emit_array_new_fixed(1, 5, 0);
        finish(&mut c);
        out.push(c);
        for (name, op) in [
            ("packed_get_u", Op::ARRAY_GET_U),
            ("packed_get_s", Op::ARRAY_GET_S),
        ] {
            let mut c = func(name, 2);
            get(&mut c, 0);
            get(&mut c, 1);
            c.emit_op_u16(op, 1, 0);
            finish(&mut c);
            out.push(c);
        }
        let mut c = func("packed_set", 3);
        for i in 0..3 {
            get(&mut c, i);
        }
        c.emit_op(Op::ARRAY_SET, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        finish(&mut c);
        out.push(c);
        let mut c = func("packed_length", 1);
        get(&mut c, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
        finish(&mut c);
        out.push(c);
        let mut c = func("into", 3);
        for i in 0..3 {
            get(&mut c, i);
        }
        call(&mut c, "intoCharCodeArray", 3);
        finish(&mut c);
        out.push(c);
        let mut c = func("from", 3);
        for i in 0..3 {
            get(&mut c, i);
        }
        call(&mut c, "fromCharCodeArray", 3);
        finish(&mut c);
        out.push(c);
        let mut c = func("generic_new", 1);
        get(&mut c, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
        finish(&mut c);
        out.push(c);
        let mut c = func("generic_get", 2);
        get(&mut c, 0);
        get(&mut c, 1);
        c.emit_op(Op::ARRAY_GET, 0);
        finish(&mut c);
        out.push(c);
        let mut c = func("generic_set", 3);
        for i in 0..3 {
            get(&mut c, i);
        }
        c.emit_op(Op::ARRAY_SET, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        finish(&mut c);
        out.push(c);
    }
    // Actual compiler ownership: only the root chunk owns module declarations.
    out[0].types = types;
    vybe_compiler::primitives::globals::declare_free_globals(&mut out);
    vybe_compiler::primitives::globals::normalize_global_table(&mut out);
    out
}
fn validate(bytes: &[u8]) -> Result<(), String> {
    match wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(bytes)
    {
        Ok(_) => Ok(()),
        Err(e) => Err(e.to_string()),
    }
}
#[test]
fn actual_import_signatures_use_own_mutable_i16_rec_group_and_unsigned_indices() {
    for classes in [0, 40] {
        let bytes = vybe_platform_wasm::write_wasm(&module(false, classes));
        let mut types = Vec::new();
        let mut singleton = Vec::new();
        let mut imports = std::collections::HashMap::new();
        for p in wasmparser::Parser::new(0).parse_all(&bytes) {
            match p.unwrap() {
                wasmparser::Payload::TypeSection(section) => {
                    for group in section {
                        let group = group.unwrap();
                        let one = group.types().len() == 1;
                        for ty in group.into_types() {
                            singleton.push(one);
                            types.push(ty);
                        }
                    }
                }
                wasmparser::Payload::ImportSection(section) => {
                    for import in section.into_imports() {
                        let import = import.unwrap();
                        if import.module == "wasm:js-string" {
                            if let wasmparser::TypeRef::Func(i) = import.ty {
                                imports.insert(import.name.to_owned(), i as usize);
                            }
                        }
                    }
                }
                _ => (),
            }
        }
        let signature = |name: &str| match &types[imports[name]].composite_type.inner {
            wasmparser::CompositeInnerType::Func(f) => f,
            _ => panic!("builtin function signature"),
        };
        let from = signature("fromCharCodeArray");
        let into = signature("intoCharCodeArray");
        let wasmparser::ValType::Ref(array_ref) = from.params()[0] else {
            panic!("real typed array ref required")
        };
        assert!(array_ref.is_nullable());
        let array_index = array_ref.type_index().unwrap().as_module_index().unwrap() as usize;
        assert!(singleton[array_index], "array must be in its own rec group");
        let wasmparser::CompositeInnerType::Array(array) = &types[array_index].composite_type.inner
        else {
            panic!("array storage required")
        };
        assert!(array.0.mutable);
        assert_eq!(array.0.element_type, wasmparser::StorageType::I16);
        assert_eq!(&from.params()[1..], &[wasmparser::ValType::I32; 2]);
        assert_eq!(
            into.params(),
            &[
                wasmparser::ValType::EXTERNREF,
                from.params()[0],
                wasmparser::ValType::I32
            ]
        );
        assert_eq!(into.results(), &[wasmparser::ValType::I32]);
        let wasmparser::ValType::Ref(result) = from.results()[0] else {
            panic!("String result ref required")
        };
        assert!(!result.is_nullable());
        assert_eq!(
            result.heap_type(),
            wasmparser::RefType::EXTERNREF.heap_type()
        );
        if classes == 40 {
            assert!(
                array_index >= 64,
                "exercise signed multi-byte heap-type LEB"
            );
        }
    }
}
fn leb(mut n: usize) -> Vec<u8> {
    let mut o = vec![];
    loop {
        let b = (n & 127) as u8;
        n >>= 7;
        o.push(if n == 0 { b } else { b | 128 });
        if n == 0 {
            return o;
        }
    }
}
fn readleb(b: &[u8], i: &mut usize) -> usize {
    let (mut n, mut shift) = (0, 0);
    loop {
        let x = b[*i];
        *i += 1;
        n |= usize::from(x & 127) << shift;
        if x & 128 == 0 {
            return n;
        }
        shift += 7;
    }
}
fn sections(bytes: &[u8]) -> Vec<(u8, std::ops::Range<usize>, usize)> {
    let mut i = 8;
    let mut o = vec![];
    while i < bytes.len() {
        let a = i;
        let id = bytes[i];
        i += 1;
        let n = readleb(bytes, &mut i);
        let start = i;
        i += n;
        assert!(i <= bytes.len());
        o.push((id, a..i, start));
    }
    o
}
fn expose(bytes: &[u8], chunks: &[Chunk]) -> Vec<u8> {
    let mut imports = 0;
    let mut func_types = vec![];
    let mut types = vec![];
    for p in wasmparser::Parser::new(0).parse_all(bytes) {
        match p.unwrap() {
            wasmparser::Payload::ImportSection(s) => {
                for i in s.into_imports() {
                    if matches!(
                        i.unwrap().ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        imports += 1;
                    }
                }
            }
            wasmparser::Payload::FunctionSection(s) => {
                func_types.extend(s.into_iter().map(Result::unwrap))
            }
            wasmparser::Payload::TypeSection(s) => {
                for g in s {
                    for ty in g.unwrap().into_types() {
                        types.push(match ty.composite_type.inner {
                            wasmparser::CompositeInnerType::Func(f) => {
                                Some((f.params().to_vec(), f.results().to_vec()))
                            }
                            _ => None,
                        });
                    }
                }
            }
            _ => (),
        }
    }
    assert!(func_types.len() >= chunks.len());
    for (i, c) in chunks.iter().enumerate() {
        let (p, r) = types[func_types[i] as usize].as_ref().unwrap();
        assert_eq!(
            p,
            &vec![wasmparser::ValType::EXTERNREF; c.arity as usize + 1]
        );
        assert_eq!(r, &[wasmparser::ValType::EXTERNREF]);
    }
    let before = sections(bytes);
    let mut out = bytes[..8].to_vec();
    for (id, range, start) in &before {
        if *id != 7 {
            out.extend_from_slice(&bytes[range.clone()]);
            continue;
        }
        let mut i = *start;
        let count = readleb(bytes, &mut i);
        let mut payload = leb(count + chunks.len() - 1);
        payload.extend_from_slice(&bytes[i..range.end]);
        for (index, c) in chunks.iter().enumerate().skip(1) {
            payload.extend(leb(c.name.len()));
            payload.extend(c.name.as_bytes());
            payload.push(0);
            payload.extend(leb(imports + index));
        }
        out.push(7);
        out.extend(leb(payload.len()));
        out.extend(payload);
    }
    let after = sections(&out);
    assert_eq!(before.len(), after.len());
    for ((id, a, _), (other, b, _)) in before.iter().zip(after) {
        assert_eq!(*id, other);
        if *id != 7 {
            assert_eq!(&bytes[a.clone()], &out[b]);
        }
    }
    validate(&out).unwrap();
    out
}

#[test]
fn declared_packed_storage_and_access_modules_validate() {
    for access in [false, true] {
        for classes in [0, 40] {
            validate(&vybe_platform_wasm::write_wasm(&module(access, classes))).unwrap();
        }
    }
}
#[test]
fn array_rows_with_class_features_keep_generic_allocations() {
    for shape in ["ordinary", "constructor", "interface"] {
        let mut root = func("owner", 0);
        root.emit_i32_const(2, 0);
        root.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 1, 0);
        finish(&mut root);
        let mut row = TypeEntry {
            name: "Characters".into(),
            kind: CompositeKind::Array,
            fields: vec!["i16".into()],
            ..Default::default()
        };
        if shape == "constructor" {
            row.constructor_chunk = Some(1);
        }
        if shape == "interface" {
            row.implements = vec![2];
        }
        root.types.push(row);
        if shape == "interface" {
            root.types.push(TypeEntry {
                name: "Contract".into(),
                is_interface: true,
                ..Default::default()
            });
        }
        let mut ctor = func("constructor", 0);
        ctor.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        finish(&mut ctor);
        let native = vybe_platform_wasm::write_wasm(&[root, ctor]);
        validate(&native).unwrap();
        let mut types = vec![];
        let mut allocation = None;
        for payload in wasmparser::Parser::new(0).parse_all(&native) {
            match payload.unwrap() {
                wasmparser::Payload::TypeSection(section) => {
                    for group in section {
                        types.extend(group.unwrap().into_types());
                    }
                }
                wasmparser::Payload::CodeSectionEntry(body) => {
                    let mut ops = body.get_operators_reader().unwrap();
                    while !ops.eof() {
                        if let wasmparser::Operator::ArrayNewDefault { array_type_index } =
                            ops.read().unwrap()
                        {
                            allocation = Some(array_type_index);
                        }
                    }
                }
                _ => (),
            }
        }
        let array = match &types[allocation.unwrap() as usize].composite_type.inner {
            wasmparser::CompositeInnerType::Array(array) => array,
            _ => panic!("allocation targets array"),
        };
        assert_eq!(
            array.0.element_type,
            if shape == "ordinary" {
                wasmparser::StorageType::I16
            } else {
                wasmparser::StorageType::Val(wasmparser::ValType::EXTERNREF)
            }
        );
    }
}
#[test]
#[ignore = "Export complete unchanged bodies for genuine V8 execution"]
fn export_native_modules() {
    let dir = std::path::PathBuf::from(std::env::var_os("VYBE_PACKED_I16_ARTIFACT_DIR").expect("set explicit native artifact directory for this ignored export gate"));
    std::fs::create_dir_all(&dir).unwrap();
    for (access, label) in [(false, "roundtrip"), (true, "packed")] {
        let chunks = module(access, 0);
        let bytes = expose(&vybe_platform_wasm::write_wasm(&chunks), &chunks);
        let mut inventory = vec![];
        for p in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::ImportSection(section) = p.unwrap() {
                for import in section.into_imports() {
                    let i = import.unwrap();
                    inventory.push(serde_json::json!({"module":i.module,"name":i.name,"mutable":match i.ty{wasmparser::TypeRef::Global(g)=>Some(g.mutable),_=>None},"kind":match i.ty{wasmparser::TypeRef::Func(_)|wasmparser::TypeRef::FuncExact(_)=>"function",wasmparser::TypeRef::Global(_)=>"global",_=>panic!("unexpected import")}}));
                }
            }
        }
        std::fs::write(dir.join(format!("candidate-{label}.wasm")), bytes).unwrap();
        std::fs::write(
            dir.join(format!("candidate-{label}.json")),
            serde_json::to_vec_pretty(&inventory).unwrap(),
        )
        .unwrap();
    }
}
