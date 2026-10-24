//! Native diagnostic exports shared by primitive regression tests.
//! Keep this source with the tests; generated artifacts may be deleted freely.
use std::ops::Range;
fn leb(mut n: usize) -> Vec<u8> {
    let mut out = vec![];
    loop {
        let b = (n & 127) as u8;
        n >>= 7;
        out.push(if n == 0 { b } else { b | 128 });
        if n == 0 {
            return out;
        }
    }
}
fn read(b: &[u8], p: &mut usize) -> usize {
    let (mut n, mut shift) = (0, 0);
    loop {
        let x = b[*p];
        *p += 1;
        n |= usize::from(x & 127) << shift;
        if x & 128 == 0 {
            return n;
        }
        shift += 7;
        assert!(shift < usize::BITS);
    }
}
fn sections(b: &[u8]) -> Vec<(u8, Range<usize>, usize)> {
    let mut p = 8;
    let mut rows = vec![];
    while p < b.len() {
        let start = p;
        let id = b[p];
        p += 1;
        let n = read(b, &mut p);
        let payload = p;
        p += n;
        assert!(p <= b.len());
        rows.push((id, start..p, payload));
    }
    rows
}
fn validate(bytes: &[u8]) {
    vybe_platform_wasm::validate_binary(bytes).expect("complete primitive native module");
}
pub fn export_chunks(chunks: &[vybe_runtime::Chunk], side: &str) {
    export_chunks_with_validator(chunks, side, validate);
}
pub fn export_chunks_with_validator(
    chunks: &[vybe_runtime::Chunk],
    side: &str,
    validate: fn(&[u8]),
) {
    assert!(
        !side.is_empty()
            && side
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || matches!(c, '-' | '_')),
        "artifact name must be one path component"
    );
    let original = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: true,
            include_vybe_metadata: false,
        },
    );
    validate(&original);
    let mut imported_functions = 0;
    let mut inventory = vec![];
    let mut function_types = vec![];
    use vybe_platform_wasm::inspection::{CompositeType, CoreModule, ImportType};
    let module = CoreModule::new(&original).unwrap();
    let types = module.types().unwrap();
    let mut defined_tags = 0;
    module.imports(|module, name, ty| {
        let detail = match ty {
            ImportType::Func { type_index, .. } => {
                imported_functions += 1;
                serde_json::json!({"kind":"function","type":type_index})
            }
            ImportType::Global(global) => {
                serde_json::json!({"kind":"global","mutable":global.mutable,"type":format!("{:?}",global.value_type)})
            }
            ImportType::Tag(type_index) => serde_json::json!({"kind":"tag","type":type_index}),
            other => panic!("unsupported diagnostic import: {other:?}"),
        };
        inventory.push(serde_json::json!({"module":module,"name":name,"detail":detail}));
        Ok(())
    }).unwrap();
    module
        .defined_functions(|index| {
            function_types.push(index);
            Ok(())
        })
        .unwrap();
    module
        .defined_tags(|_| {
            defined_tags += 1;
            Ok(())
        })
        .unwrap();
    assert!(function_types.len() >= chunks.len());
    let mut names = std::collections::BTreeMap::new();
    for (id, range, payload) in sections(&original) {
        if id != 0 {
            continue;
        }
        let mut p = payload;
        let n = read(&original, &mut p);
        let tag = &original[p..p + n];
        p += n;
        if tag != b"name" {
            continue;
        }
        while p < range.end {
            let sub = original[p];
            p += 1;
            let n = read(&original, &mut p);
            let end = p + n;
            if sub == 1 {
                let count = read(&original, &mut p);
                for _ in 0..count {
                    let index = read(&original, &mut p);
                    let n = read(&original, &mut p);
                    names.insert(
                        index,
                        std::str::from_utf8(&original[p..p + n]).unwrap().to_owned(),
                    );
                    p += n;
                }
            }
            p = end;
        }
    }
    for (index, chunk) in chunks.iter().enumerate() {
        assert_eq!(
            names.get(&(imported_functions + index)),
            Some(&chunk.name),
            "writer function index provenance"
        );
    }
    let mut bytes = original[..8].to_vec();
    let mut exports = vec![];
    let mut aliases = vec![];
    let mut export_section = false;
    for (id, range, payload) in sections(&original) {
        if id != 7 {
            bytes.extend_from_slice(&original[range]);
            continue;
        }
        export_section = true;
        let mut p = payload;
        let count = read(&original, &mut p);
        let mut data = leb(count + chunks.len() + defined_tags);
        data.extend_from_slice(&original[p..range.end]);
        for (index, chunk) in chunks.iter().enumerate() {
            let name = format!("guest_chunk_{index}");
            data.extend(leb(name.len()));
            data.extend(name.bytes());
            data.push(0);
            data.extend(leb(imported_functions + index));
            let CompositeType::Func { params, results } =
                &types[function_types[index] as usize].definition().composite
            else {
                panic!("exported chunk type is not a function");
            };
            exports.push(serde_json::json!({"name":chunk.name,"export":name,"chunk":index,"params":params.iter().map(|t|format!("{t:?}")).collect::<Vec<_>>(),"results":results.iter().map(|t|format!("{t:?}")).collect::<Vec<_>>() }));
        }
        for tag in 0..defined_tags {
            let name = format!("guest_exception_tag_{tag}");
            data.extend(leb(name.len()));
            data.extend(name.bytes());
            data.push(4);
            data.extend(leb(tag));
        }
        bytes.push(7);
        bytes.extend(leb(data.len()));
        bytes.extend(data);
    }
    assert!(
        export_section,
        "writer must provide its normal export section"
    );
    let before = sections(&original);
    let after = sections(&bytes);
    assert_eq!(before.len(), after.len());
    for ((a, x, _), (b, y, _)) in before.iter().zip(after.iter()) {
        assert_eq!(a, b);
        if *a != 7 {
            assert_eq!(&original[x.clone()], &bytes[y.clone()]);
        }
    }
    validate(&bytes);
    for c in chunks {
        for init in &c.global_inits {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = &init.init {
                aliases.push(serde_json::json!({"name":init.name,"chunk":index,"chunkName":chunks[*index].name,"embeddingValue":index}));
            }
        }
    }
    let root = std::env::var_os("VYBE_PRIMITIVE_ARTIFACT_DIR")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| std::env::temp_dir().join("vybe-primitive-native-artifacts"));
    let dir = root.join(side);
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::write(dir.join("complete.wasm"), bytes).unwrap();
    std::fs::write(dir.join("complete.json"),serde_json::to_vec_pretty(&serde_json::json!({"inventory":inventory,"exports":exports,"aliases":aliases,"scope":"complete writer module; only export section extended; standard builtin/GC operations and explicit RefFunc embedding initialization"})).unwrap()).unwrap();
    println!("{}", dir.display());
}
