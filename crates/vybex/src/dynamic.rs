use std::path::{Path, PathBuf};

use vybe_bytecode::chunk::Chunk;
use vybe_bytecode::value::{Function, Object, ObjectKind};
use vybe_bytecode::{VM, Value};
use vybe_compiler::bundle::{Bundle, CompiledBundle, EntryPoint, SourceFile};
use vybe_compiler::compiler::HostImportMetadata;
use vybe_compiler::languages::{self, Language};
use vybe_host::{Capabilities, Capability};

#[derive(Debug)]
pub struct DynamicCompilation {
    pub chunks: Vec<Chunk>,
    pub host_imports: HostImportMetadata,
}

pub struct RuntimeCompilerService<'vm> {
    vm: &'vm mut VM,
    caps: Capabilities,
}

impl<'vm> RuntimeCompilerService<'vm> {
    pub fn new(vm: &'vm mut VM) -> Self {
        Self::with_capabilities(vm, Capabilities::all())
    }

    pub fn with_capabilities(vm: &'vm mut VM, caps: Capabilities) -> Self {
        Self { vm, caps }
    }

    pub fn vm(&mut self) -> &mut VM {
        self.vm
    }

    pub fn compile_bundle(&mut self, bundle: &Bundle) -> Result<DynamicCompilation, String> {
        let compiled = bundle.compile_full_with_modules(&self.vm.modules)?;
        Ok(DynamicCompilation {
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
        })
    }

    pub fn compile_path(&mut self, path: &Path) -> Result<DynamicCompilation, String> {
        let bundle = vybe_compiler::projects::load(path)?;
        self.compile_bundle(&bundle)
    }

    pub fn compile_source(
        &mut self,
        source: impl Into<String>,
        language: Language,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<DynamicCompilation, String> {
        self.ensure_dynamic_compile_allowed()?;
        let bundle = bundle_from_source(source, language, virtual_path);
        self.compile_bundle(&bundle)
    }

    pub fn compile_source_by_name(
        &mut self,
        source: impl Into<String>,
        language_name: &str,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<DynamicCompilation, String> {
        let language = languages::find_by_name(language_name)
            .ok_or_else(|| format!("unknown language: {language_name}"))?;
        self.compile_source(source, language, virtual_path)
    }

    pub fn can_dynamic_compile(&self) -> bool {
        self.caps.has(Capability::DynamicCompile)
    }

    fn ensure_dynamic_compile_allowed(&self) -> Result<(), String> {
        if self.can_dynamic_compile() {
            return Ok(());
        }
        Err("Dynamic compilation is disabled by the current capability set (missing Capability::DynamicCompile)".to_string())
    }

    pub fn run_compiled(&mut self, compiled: DynamicCompilation) -> Result<Value, String> {
        let base_chunk_index = self.vm.chunks.len();
        crate::host_imports::install(self.vm, &compiled.host_imports);
        install_chunk_globals(self.vm, &compiled.chunks, base_chunk_index);
        self.vm.run(compiled.chunks).map_err(|e| e.to_string())
    }

    pub fn compile_and_run_bundle(&mut self, bundle: &Bundle) -> Result<Value, String> {
        let compiled = self.compile_bundle(bundle)?;
        self.run_compiled(compiled)
    }

    pub fn compile_and_run_path(&mut self, path: &Path) -> Result<Value, String> {
        let compiled = self.compile_path(path)?;
        self.run_compiled(compiled)
    }

    pub fn compile_and_run_source(
        &mut self,
        source: impl Into<String>,
        language: Language,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<Value, String> {
        let compiled = self.compile_source(source, language, virtual_path)?;
        self.run_compiled(compiled)
    }
}

pub fn bundle_from_source(
    source: impl Into<String>,
    language: Language,
    virtual_path: impl Into<PathBuf>,
) -> Bundle {
    let path = virtual_path.into();
    let name = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("dynamic")
        .to_string();

    Bundle {
        name,
        language,
        sources: vec![SourceFile {
            path,
            code: source.into(),
        }],
        wasm_files: Vec::new(),
        entry_point: EntryPoint::Auto,
    }
}

pub fn language_for_path(path: &Path) -> Option<Language> {
    let ext = path.extension()?.to_str()?;
    languages::find_by_extension(ext)
}

pub fn install_chunk_globals(vm: &mut VM, chunks: &[Chunk], base_chunk_index: usize) {
    use std::sync::{Arc, Mutex};

    for (idx, chunk) in chunks.iter().enumerate() {
        if !should_publish_chunk_name(&chunk.name) {
            continue;
        }

        let func = Function {
            name: Some(chunk.name.clone()),
            arity: chunk.arity,
            chunk_index: base_chunk_index + idx,
            upvalues: vec![],
        };
        let mut obj = Object::new();
        obj.kind = ObjectKind::Function(func);
        let val = Value::Object(Arc::new(Mutex::new(obj)));
        vm.globals.insert(chunk.name.to_lowercase(), val);
    }
}

fn should_publish_chunk_name(name: &str) -> bool {
    !name.is_empty()
        && name != "<script>"
        && name != "<bootstrap>"
        && !name.starts_with("__stdlib_")
}

pub fn into_dynamic_compilation(compiled: CompiledBundle) -> DynamicCompilation {
    DynamicCompilation {
        chunks: compiled.chunks,
        host_imports: compiled.host_imports,
    }
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use super::RuntimeCompilerService;
    use vybe_bytecode::{VM, Value};

    struct DynamicSmokeCase {
        language: &'static str,
        virtual_path: &'static str,
        source: &'static str,
    }

    fn configured_vm() -> VM {
        let mut vm = VM::new();
        let _gui = vybe_host::register_all_with_gui(&mut vm);
        vybe_host::setup_namespaces(&mut vm);
        vm
    }

    fn assert_numeric_value(value: Value, expected: f64) {
        match value {
            Value::I32(n) => assert_eq!(n as f64, expected),
            Value::I64(n) => assert_eq!(n as f64, expected),
            Value::F64(n) => assert_eq!(n, expected),
            other => panic!("expected numeric value, got {other:?}"),
        }
    }

    #[test]
    fn compiles_source_into_live_vm_and_publishes_function_globals() {
        let mut vm = configured_vm();

        {
            let mut service = RuntimeCompilerService::new(&mut vm);
            service
                .compile_and_run_source(
                    "function greet() { return 7; }",
                    vybe_compiler::languages::find_by_name("js").expect("js language"),
                    PathBuf::from("dynamic/greet.js"),
                )
                .expect("compile and run greet");

            service
                .compile_and_run_source(
                    "function callGreet() { return greet(); }",
                    vybe_compiler::languages::find_by_name("js").expect("js language"),
                    PathBuf::from("dynamic/call_greet.js"),
                )
                .expect("compile and run callGreet");
        }

        let greet = vm.globals.get("greet").cloned().expect("greet global");
        let call_greet = vm.globals.get("callgreet").cloned().expect("callGreet global");

        assert_numeric_value(vm.invoke(&greet, &[]).expect("invoke greet"), 7.0);
        assert_numeric_value(vm.invoke(&call_greet, &[]).expect("invoke callGreet"), 7.0);
    }

    #[test]
    fn dynamic_execution_smoke_matrix_for_supported_languages() {
        let cases = [
            DynamicSmokeCase {
                language: "js",
                virtual_path: "dynamic/matrix.js",
                source: "let x = 7;",
            },
            DynamicSmokeCase {
                language: "php",
                virtual_path: "dynamic/matrix.php",
                source: "<?php $x = 7;",
            },
            DynamicSmokeCase {
                language: "python",
                virtual_path: "dynamic/matrix.py",
                source: "x = 7",
            },
            DynamicSmokeCase {
                language: "ruby",
                virtual_path: "dynamic/matrix.rb",
                source: "x = 7",
            },
            DynamicSmokeCase {
                language: "dart",
                virtual_path: "dynamic/matrix.dart",
                source: "var x = 7;",
            },
            DynamicSmokeCase {
                language: "vb",
                virtual_path: "dynamic/matrix.vb",
                source: "Dim x As Integer = 7",
            },
            DynamicSmokeCase {
                language: "csharp",
                virtual_path: "dynamic/matrix.cs",
                source: "int x = 7;",
            },
            DynamicSmokeCase {
                language: "pascal",
                virtual_path: "dynamic/matrix.pas",
                source: "program T; var x: Integer; begin x := 7; end.",
            },
            DynamicSmokeCase {
                language: "cobol",
                virtual_path: "dynamic/matrix.cob",
                source: "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 7.\nPROCEDURE DIVISION.\n    STOP RUN.",
            },
            DynamicSmokeCase {
                language: "fortran",
                virtual_path: "dynamic/matrix.f90",
                source: "program test\n  integer :: x\n  x = 7\nend program test",
            },
        ];

        for case in cases {
            let mut vm = configured_vm();
            let mut service = RuntimeCompilerService::new(&mut vm);
            let compiled = service
                .compile_source_by_name(
                    case.source,
                    case.language,
                    PathBuf::from(case.virtual_path),
                )
                .unwrap_or_else(|err| panic!("{} dynamic compile failed: {}", case.language, err));

            service
                .run_compiled(compiled)
                .unwrap_or_else(|err| panic!("{} dynamic run failed: {}", case.language, err));
        }
    }

    #[test]
    fn dynamic_compile_requires_capability_for_source_text() {
        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::with_capabilities(&mut vm, vybe_host::Capabilities::safe());

        let err = service
            .compile_source_by_name("let x = 7;", "js", PathBuf::from("dynamic/locked.js"))
            .expect_err("dynamic compile should be denied without capability");

        assert!(err.contains("DynamicCompile"), "unexpected error: {err}");
    }
}