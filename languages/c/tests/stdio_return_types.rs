use std::path::Path;
use vybe_ast::{BinOp, BitLane, ExprKind, IntOp, StmtKind};
use vybe_runtime::VM;

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

#[test]
fn generated_file_helpers_declare_numeric_results_not_handle_parameters() {
    let source = std::fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup/static_file_registry.c"),
    )
    .unwrap();
    let module = vybe_language_c::parse(&source).unwrap();
    for (expected_name, expected_type) in [
        ("__c_remove_h", "int"),
        ("__c_rename_h", "int"),
        ("__c_fsync_h", "int"),
        ("__c_fclose_h", "int"),
        ("__c_fgetc_h", "int"),
        ("__c_ungetc_h", "int"),
        ("__c_ftell_h", "long"),
        ("__c_rand_h", "int"),
        ("__c_fwrite_h", "uint64"),
    ] {
        let (params, result) = module
            .body
            .iter()
            .find_map(|statement| {
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    return_type,
                    ..
                } = &statement.kind
                    && name == expected_name
                {
                    Some((params, return_type))
                } else {
                    None
                }
            })
            .unwrap_or_else(|| panic!("missing {expected_name}"));
        assert_eq!(result.as_deref(), Some(expected_type), "{expected_name}");
        assert!(
            params.iter().all(|parameter| parameter.type_hint.is_none()),
            "{expected_name}: buffer/handle representation must remain hybrid"
        );
    }
    for name in ["__c_fread_h", "__c_fopen_h"] {
        let result = module
            .body
            .iter()
            .find_map(|statement| {
                if let StmtKind::FunctionDecl {
                    name: candidate,
                    return_type,
                    ..
                } = &statement.kind
                    && candidate == name
                {
                    Some(return_type)
                } else {
                    None
                }
            })
            .unwrap();
        assert!(
            result.is_none(),
            "{name}: buffer/handle result is not numeric"
        );
    }
}

#[test]
fn rewritten_stdio_retains_status_and_size_result_lanes() {
    let module = vybe_language_c::parse(
        "#include <stdio.h>\n#include <stdlib.h>\n
         int probe(FILE *f, unsigned char *p) {
             return fseek(f,0,0)!=0 || fseeko(f,0,0)==-1 || fread(p,1,16,f)!=16;
         }",
    )
    .unwrap();
    let mut declarations = 0;
    for statement in &module.body {
        if let StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            ..
        } = &statement.kind
        {
            let expected = match name.as_str() {
                "__c_path_present" | "__c_fseek_h" => Some("int"),
                "__c_fread_into_h"
                | "__c_fread_into_linear_h"
                | "__c_fread_into_array_h"
                | "__c_fread_into_struct_array_h" => Some("uint64"),
                _ => None,
            };
            if let Some(expected) = expected {
                assert_eq!(return_type.as_deref(), Some(expected), "{name}");
                assert!(
                    params.iter().all(|param| param.type_hint.is_none()),
                    "{name}"
                );
                declarations += 1;
            }
            if name == "__c_fread_h" {
                assert!(return_type.is_none(), "byte-array helper is not a count");
            }
        }
    }
    assert!(declarations >= 2);
    let statement = module.body.iter().find(|statement|
        matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "probe")).unwrap();
    let mut status = 0;
    let mut sizes = 0;
    statement.walk_exprs(&mut |value| {
        if let ExprKind::Binary {
            op: BinOp::Integer(op, lane),
            left,
            ..
        } = &value.kind
            && matches!(op, IntOp::Eq | IntOp::Ne)
            && let ExprKind::Cast { type_name, expr } = &left.kind
        {
            let mut helpers = Vec::new();
            expr.walk_exprs(&mut |value| {
                if let ExprKind::Call { callee, .. } = &value.kind
                    && let ExprKind::Ident(name) = &callee.kind
                {
                    helpers.push(name.clone());
                }
            });
            if helpers.iter().any(|name| name == "__c_fseek_h") {
                assert_eq!(type_name, "int");
                assert_eq!(*lane, BitLane::W32);
                status += 1;
            }
            if helpers
                .iter()
                .any(|name| name.starts_with("__c_fread_into_"))
            {
                assert_eq!(type_name, "uint64");
                assert_eq!(*lane, BitLane::W64);
                sizes += 1;
            }
        }
    });
    assert_eq!((status, sizes), (2, 1));
}

#[test]
fn rewritten_fread_uses_the_declared_size_width_and_leaves_macros_alone() {
    for (declaration, result, expected) in [
        ("", "unsigned int", "uint32"),
        ("", "unsigned long long", "uint64"),
        (
            "typedef unsigned int read_count_t;",
            "read_count_t",
            "uint32",
        ),
    ] {
        let module = vybe_language_c::parse(&format!(
            "typedef int FILE; {declaration}
             {result} fread(void *, unsigned long long, unsigned long long, FILE *);
             int probe(FILE *f, unsigned char *p) {{ return fread(p,1,16,f)!=16; }}"
        ))
        .unwrap();
        let mut casts = Vec::new();
        for statement in &module.body {
            if matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "probe") {
                statement.walk_exprs(&mut |value| {
                    if let ExprKind::Cast { type_name, .. } = &value.kind {
                        casts.push(type_name.clone());
                    }
                });
            }
        }
        assert!(
            casts.iter().any(|name| name == expected),
            "missing {expected} fread result: {casts:?}"
        );
        for statement in &module.body {
            if let StmtKind::FunctionDecl {
                name, return_type, ..
            } = &statement.kind
                && name.starts_with("__c_fread_into_")
            {
                assert_eq!(return_type.as_deref(), Some(expected), "{name}");
            }
        }
    }
    let module = vybe_language_c::parse(
        "#define fread(p,s,n,f) 0.5\n double probe(void) { return fread(0,0,0,0); }",
    )
    .unwrap();
    let statement = module.body.iter().find(|statement|
        matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "probe")).unwrap();
    statement.walk_exprs(&mut |value| {
        assert!(!matches!(&value.kind, ExprKind::Cast { type_name, .. }
            if matches!(type_name.as_str(), "uint32" | "uint64")));
    });
}

#[test]
fn rewritten_stdio_hot_copy_has_no_dynamic_equality_in_main() {
    struct RestoreDirectory(std::path::PathBuf);
    impl Drop for RestoreDirectory {
        fn drop(&mut self) {
            std::env::set_current_dir(&self.0).unwrap();
        }
    }
    let _restore = RestoreDirectory(std::env::current_dir().unwrap());
    std::env::set_current_dir(Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")).unwrap();
    for fixture in ["managed_linear_copy_hot", "typed_stdio_return_contracts", "stdio_eof_integer_macro"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        let mut vm = runtime();
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path)
            .unwrap();
        for name in ["main", "read_buffer", "is_eof"] {
            if let Some(chunk) = compiled.chunks.iter().find(|chunk| chunk.name == name) {
                let dump = vybe_runtime::debug::disassemble(chunk);
                eprintln!(
                    "{fixture}/{name}: dynamic equality sites={}",
                    dump.matches("__vybe_dyneq").count()
                );
                assert!(
                    !dump.contains("__vybe_dyneq"),
                    "{fixture}/{name}: dynamic equality\n{dump}"
                );
                if let Some(dir) = std::env::var_os("C_STDIO_WASM_DIR") {
                    std::fs::write(
                        Path::new(&dir).join(format!("typed-stdio-{fixture}-{name}.dump")),
                        dump,
                    )
                    .unwrap();
                }
            }
        }
        let wasm = vybe_platform_wasm::write_wasm_with_options(
            &compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm)
            .unwrap();
        if let Some(dir) = std::env::var_os("C_STDIO_WASM_DIR") {
            std::fs::write(
                Path::new(&dir).join(format!("typed-stdio-{fixture}.wasm")),
                wasm,
            )
            .unwrap();
        }
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled)
            .unwrap();
        assert_eq!(vm.pending_exit_code, 0, "{fixture}");
    }
}
