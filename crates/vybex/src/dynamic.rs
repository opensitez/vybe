use std::cell::RefCell;
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};

use vybe_bytecode::chunk::Import;
use vybe_bytecode::opcode::{Op, OperandFormat};
use vybe_bytecode::chunk::Chunk;
use vybe_bytecode::value::{Function, Object, ObjectKind};
use vybe_bytecode::{ImportTarget, VM, Value};
use vybe_compiler::bundle::{Bundle, CompiledBundle, EntryPoint, SourceFile};
use vybe_compiler::compiler::HostImportMetadata;
use vybe_compiler::languages::{self, Language};
use vybe_host::{Capabilities, Capability};

thread_local! {
    static ACTIVE_PHP_RUNTIME: RefCell<Option<*mut PhpIncludeRuntime>> = const { RefCell::new(None) };
}

#[derive(Debug)]
pub struct DynamicCompilation {
    pub chunks: Vec<Chunk>,
    pub host_imports: HostImportMetadata,
    pub entry_path: Option<PathBuf>,
}

pub struct RuntimeCompilerService<'vm> {
    vm: &'vm mut VM,
    caps: Capabilities,
    php_runtime: PhpIncludeRuntime,
}

struct PhpIncludeRuntime {
    caps: Capabilities,
    vm: *mut VM,
    current_paths: Vec<PathBuf>,
    included_once: HashSet<PathBuf>,
    active_imports: Vec<Import>,
    active_resolved_imports: Vec<ImportTarget>,
}

struct ActivePhpRuntimeGuard {
    previous: Option<*mut PhpIncludeRuntime>,
}

impl<'vm> RuntimeCompilerService<'vm> {
    pub fn new(vm: &'vm mut VM) -> Self {
        Self::with_capabilities(vm, Capabilities::all())
    }

    pub fn with_capabilities(vm: &'vm mut VM, caps: Capabilities) -> Self {
        Self {
            vm,
            caps: caps.clone(),
            php_runtime: PhpIncludeRuntime::new(caps),
        }
    }

    pub fn vm(&mut self) -> &mut VM {
        self.vm
    }

    pub fn compile_bundle(&mut self, bundle: &Bundle) -> Result<DynamicCompilation, String> {
        ensure_php_runtime_registered(self.vm);
        let compiled = bundle.compile_full_with_modules(&self.vm.modules)?;
        Ok(DynamicCompilation {
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
            entry_path: bundle.sources.first().map(|source| source.path.clone()),
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
        ensure_php_runtime_registered(self.vm);
        let base_chunk_index = self.vm.chunks.len();
        crate::host_imports::install(self.vm, &compiled.host_imports);
        install_chunk_globals(self.vm, &compiled.chunks, base_chunk_index);
        let active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.clone())
            .unwrap_or_default();
        let active_resolved_imports = resolve_imports(self.vm, &active_imports)?;
        let _php_runtime = self.php_runtime.activate(
            self.vm,
            compiled.entry_path.as_deref(),
            active_imports,
            active_resolved_imports,
        );
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
    let source = source.into();
    let name = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("dynamic")
        .to_string();
    let code = if language.name == "php" {
        preprocess_dynamic_php_source(&source, &path)
    } else {
        source
    };

    Bundle {
        name,
        language,
        sources: vec![SourceFile {
            path,
            code,
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

fn preprocess_dynamic_php_source(source: &str, source_path: &Path) -> String {
    let absolute = absolutize_path(source_path);
    let file_literal = php_single_quoted_literal(&absolute.to_string_lossy());
    let dir_literal = php_single_quoted_literal(
        absolute
            .parent()
            .unwrap_or(Path::new("."))
            .to_string_lossy()
            .as_ref(),
    );

    replace_php_magic_constants(source, &file_literal, &dir_literal)
}

fn replace_php_magic_constants(source: &str, file_literal: &str, dir_literal: &str) -> String {
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum State {
        Normal,
        SingleQuoted,
        DoubleQuoted,
        LineComment,
        BlockComment,
    }

    let bytes = source.as_bytes();
    let mut out = String::with_capacity(source.len() + 32);
    let mut index = 0usize;
    let mut state = State::Normal;

    while index < bytes.len() {
        match state {
            State::Normal => {
                if starts_with_magic_constant(bytes, index, b"__FILE__")
                    && is_identifier_boundary(bytes, index, b"__FILE__".len())
                {
                    out.push_str(file_literal);
                    index += b"__FILE__".len();
                    continue;
                }
                if starts_with_magic_constant(bytes, index, b"__DIR__")
                    && is_identifier_boundary(bytes, index, b"__DIR__".len())
                {
                    out.push_str(dir_literal);
                    index += b"__DIR__".len();
                    continue;
                }

                if bytes[index] == b'\'' {
                    state = State::SingleQuoted;
                } else if bytes[index] == b'"' {
                    state = State::DoubleQuoted;
                } else if bytes[index] == b'/' && index + 1 < bytes.len() && bytes[index + 1] == b'/' {
                    state = State::LineComment;
                } else if bytes[index] == b'#' {
                    state = State::LineComment;
                } else if bytes[index] == b'/' && index + 1 < bytes.len() && bytes[index + 1] == b'*' {
                    state = State::BlockComment;
                }

                out.push(bytes[index] as char);
                index += 1;
            }
            State::SingleQuoted => {
                out.push(bytes[index] as char);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.push(bytes[index + 1] as char);
                    index += 2;
                    continue;
                }
                if bytes[index] == b'\'' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::DoubleQuoted => {
                out.push(bytes[index] as char);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.push(bytes[index + 1] as char);
                    index += 2;
                    continue;
                }
                if bytes[index] == b'"' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::LineComment => {
                out.push(bytes[index] as char);
                if bytes[index] == b'\n' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::BlockComment => {
                out.push(bytes[index] as char);
                if bytes[index] == b'*' && index + 1 < bytes.len() && bytes[index + 1] == b'/' {
                    out.push('/');
                    index += 2;
                    state = State::Normal;
                    continue;
                }
                index += 1;
            }
        }
    }

    out
}

fn starts_with_magic_constant(bytes: &[u8], index: usize, needle: &[u8]) -> bool {
    bytes.get(index..index + needle.len()) == Some(needle)
}

fn is_identifier_boundary(bytes: &[u8], index: usize, len: usize) -> bool {
    let before_ok = index == 0 || !is_php_identifier_byte(bytes[index - 1]);
    let after = index + len;
    let after_ok = after >= bytes.len() || !is_php_identifier_byte(bytes[after]);
    before_ok && after_ok
}

fn is_php_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

fn php_single_quoted_literal(value: &str) -> String {
    let escaped = value.replace('\\', "\\\\").replace('\'', "\\'");
    format!("'{}'", escaped)
}

fn absolutize_path(path: &Path) -> PathBuf {
    if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(path)
    }
}

pub fn into_dynamic_compilation(compiled: CompiledBundle) -> DynamicCompilation {
    DynamicCompilation {
        chunks: compiled.chunks,
        host_imports: compiled.host_imports,
        entry_path: None,
    }
}

impl PhpIncludeRuntime {
    fn new(caps: Capabilities) -> Self {
        Self {
            caps,
            vm: std::ptr::null_mut(),
            current_paths: Vec::new(),
            included_once: HashSet::new(),
            active_imports: Vec::new(),
            active_resolved_imports: Vec::new(),
        }
    }

    fn activate(
        &mut self,
        vm: &mut VM,
        entry_path: Option<&Path>,
        active_imports: Vec<Import>,
        active_resolved_imports: Vec<ImportTarget>,
    ) -> ActivePhpRuntimeGuard {
        self.vm = vm as *mut VM;
        self.current_paths.clear();
        if let Some(path) = entry_path {
            self.current_paths.push(path.to_path_buf());
        }
        self.active_imports = active_imports;
        self.active_resolved_imports = active_resolved_imports;
        let previous = ACTIVE_PHP_RUNTIME.with(|slot| slot.replace(Some(self as *mut _)));
        ActivePhpRuntimeGuard { previous }
    }

    fn handle_dynamic_include(&mut self, args: &[Value]) -> Result<Value, String> {
        if !self.caps.has(Capability::DynamicCompile) {
            return Ok(Value::Bool(false));
        }

        if self.vm.is_null() {
            return Ok(Value::Bool(false));
        }

        let vm = unsafe { &mut *self.vm };

        let kind = args.first().map(value_to_string).unwrap_or_default();
        let target = args.get(1).map(value_to_string).unwrap_or_default();
        let caller = self
            .current_paths
            .last()
            .cloned()
            .ok_or_else(|| "dynamic PHP include has no active caller path".to_string())?;
        let entry = self
            .current_paths
            .first()
            .cloned()
            .ok_or_else(|| "dynamic PHP include has no active entry path".to_string())?;
        let resolved_path = resolve_php_include_path(&entry, &caller, &target);
        let canonical_path = resolved_path.canonicalize().unwrap_or_else(|_| resolved_path.clone());

        if matches!(kind.as_str(), "include_once" | "require_once")
            && self.included_once.contains(&canonical_path)
        {
            return Ok(Value::I32(1));
        }

        let source = match fs::read_to_string(&resolved_path) {
            Ok(source) => source,
            Err(_) => return Ok(Value::Bool(false)),
        };

        let language = languages::find_by_name("php").ok_or_else(|| "php language profile missing".to_string())?;
        let bundle = bundle_from_source(source, language, resolved_path.clone());
        let mut compiled = self.compile_dynamic_php(vm, &bundle, &entry)?;

        let base_chunk_index = vm.chunks.len();
        crate::host_imports::install(vm, &compiled.host_imports);
        install_chunk_globals(vm, &compiled.chunks, base_chunk_index);

        let child_active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.clone())
            .unwrap_or_default();
        let (merged_active_imports, child_import_remap) =
            merge_imports(&self.active_imports, &child_active_imports);
        remap_import_operands(&mut compiled.chunks, &child_import_remap)?;
        let merged_active_resolved_imports = resolve_imports(vm, &merged_active_imports)?;
        let saved_active_imports = std::mem::replace(&mut self.active_imports, merged_active_imports);
        let saved_active_resolved_imports = std::mem::replace(
            &mut self.active_resolved_imports,
            merged_active_resolved_imports.clone(),
        );

        self.current_paths.push(canonical_path.clone());
        let result = vm
            .run_linked(compiled.chunks, merged_active_resolved_imports)
            .map_err(|err| err.to_string());
        self.current_paths.pop();
        self.active_imports = saved_active_imports;
        self.active_resolved_imports = saved_active_resolved_imports;

        match result {
            Ok(Value::Null) => {
                if matches!(kind.as_str(), "include_once" | "require_once") {
                    self.included_once.insert(canonical_path);
                }
                Ok(Value::I32(1))
            }
            Ok(value) => {
                if matches!(kind.as_str(), "include_once" | "require_once") {
                    self.included_once.insert(canonical_path);
                }
                Ok(value)
            }
            Err(_) => Ok(Value::Bool(false)),
        }
    }

    fn compile_dynamic_php(
        &self,
        vm: &VM,
        bundle: &Bundle,
        entry_path: &Path,
    ) -> Result<DynamicCompilation, String> {
        let compiled = bundle.compile_full_with_modules_and_php_entry_override(
            &vm.modules,
            Some(entry_path),
        )?;
        Ok(DynamicCompilation {
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
            entry_path: bundle.sources.first().map(|source| source.path.clone()),
        })
    }
}

impl Drop for ActivePhpRuntimeGuard {
    fn drop(&mut self) {
        ACTIVE_PHP_RUNTIME.with(|slot| {
            slot.replace(self.previous.take());
        });
    }
}

fn ensure_php_runtime_registered(vm: &mut VM) {
    let key = ("vybe:php".to_string(), "dynamic_include".to_string());
    if vm.host_registry.contains_key(&key) {
        return;
    }

    vm.register_host_fn("vybe:php", "dynamic_include", Box::new(|_ctx, args| {
        ACTIVE_PHP_RUNTIME.with(|slot| {
            let Some(runtime_ptr) = *slot.borrow() else {
                return Value::Bool(false);
            };
            let runtime = unsafe { &mut *runtime_ptr };
            runtime.handle_dynamic_include(args).unwrap_or(Value::Bool(false))
        })
    }));
}

fn resolve_php_include_path(entry_path: &Path, caller_path: &Path, target: &str) -> PathBuf {
    let target_path = PathBuf::from(target);
    if target_path.is_absolute() {
        return target_path;
    }
    let base_path = if is_explicit_relative_php_include(&target_path) {
        caller_path
    } else {
        entry_path
    };
    base_path
        .parent()
        .map(|parent| parent.join(target_path.clone()))
        .unwrap_or(target_path)
}

fn is_explicit_relative_php_include(path: &Path) -> bool {
    let raw = path.to_string_lossy();
    raw == "." || raw == ".." || raw.starts_with("./") || raw.starts_with("../") || raw.starts_with(".\\") || raw.starts_with("..\\")
}

fn value_to_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        _ => format!("{value}"),
    }
}

fn resolve_imports(vm: &VM, imports: &[Import]) -> Result<Vec<ImportTarget>, String> {
    let mut resolved = Vec::with_capacity(imports.len());
    for import in imports {
        let key = (import.module.clone(), import.name.clone());
        if let Some(&idx) = vm.host_registry.get(&key) {
            resolved.push(ImportTarget::Host(idx));
            continue;
        }
        if import.module == "*" {
            let candidates = [import.name.clone(), import.name.to_lowercase()];
            if let Some(global_name) = candidates.iter().find(|name| vm.globals.contains_key(name.as_str())) {
                resolved.push(ImportTarget::StdlibRedirect(global_name.clone()));
                continue;
            }
        }
        let candidates = [
            format!("__vybe_{}", import.name),
            format!("__vybe_{}", import.name.to_lowercase()),
        ];
        if let Some(global_name) = candidates.iter().find(|name| vm.globals.contains_key(name.as_str())) {
            resolved.push(ImportTarget::StdlibRedirect(global_name.clone()));
            continue;
        }
        return Err(format!(
            "Unresolved import: \"{}\" \"{}\"",
            import.module, import.name
        ));
    }
    Ok(resolved)
}

fn merge_imports(active_imports: &[Import], child_imports: &[Import]) -> (Vec<Import>, Vec<u16>) {
    let mut merged = active_imports.to_vec();
    let mut remap = Vec::with_capacity(child_imports.len());

    for child in child_imports {
        let index = merged
            .iter()
            .position(|active| active.module == child.module && active.name == child.name)
            .unwrap_or_else(|| {
                merged.push(child.clone());
                merged.len() - 1
            });
        remap.push(index as u16);
    }

    (merged, remap)
}

fn remap_import_operands(chunks: &mut [Chunk], remap: &[u16]) -> Result<(), String> {
    for chunk in chunks {
        let code = &mut chunk.code;
        let mut ip = 0;
        while ip < code.len() {
            if ip + 1 >= code.len() {
                break;
            }
            let Some(op) = Op::decode(code[ip], code[ip + 1]) else {
                ip += 2;
                continue;
            };
            if op == Op::CALL_IMPORT && ip + 3 < code.len() {
                let old_idx = ((code[ip + 2] as u16) << 8) | (code[ip + 3] as u16);
                let Some(&new_idx) = remap.get(old_idx as usize) else {
                    return Err(format!("dynamic include import remap missing entry for index {old_idx}"));
                };
                code[ip + 2] = (new_idx >> 8) as u8;
                code[ip + 3] = (new_idx & 0xff) as u8;
            }
            ip += 2;
            match op.operand_format() {
                OperandFormat::Closure => {
                    ip += 2 + 1;
                    if ip > 0 && ip - 1 < code.len() {
                        let uv_count = code[ip - 1] as usize;
                        ip += uv_count * 2;
                    }
                }
                OperandFormat::BrTable => {
                    if ip < code.len() {
                        let count = code[ip] as usize;
                        ip += 2 + count;
                    }
                }
                OperandFormat::TryTable => {
                    if ip < code.len() {
                        let count = code[ip] as usize;
                        ip += 1 + count * 3;
                    }
                }
                format => {
                    ip += format.fixed_size();
                }
            }
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;
    use std::sync::{Arc, Mutex};
    use std::time::{SystemTime, UNIX_EPOCH};

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

    #[test]
    fn php_dynamic_include_executes_expression_form() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-include-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let partial_path = base.join("partial.php");

        std::fs::write(
            &main_path,
            "<?php $path = 'partial.php'; $value = include $path; $result = get_value() + $value;",
        )
        .expect("write main php");
        std::fs::write(
            &partial_path,
            "<?php function get_value() { return 8; } return 34;",
        )
        .expect("write partial php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php with dynamic include");

        match vm.globals.get("$result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected $result global, got {other:?}"),
        }
    }

    #[test]
    fn php_dynamic_include_resolves_magic_file_constants_in_nested_runtime_paths() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-magic-include-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let child_path = base.join("child.php");
        let grandchild_path = base.join("grandchild.php");

        std::fs::write(
            &main_path,
            "<?php $path = 'child.php'; $result = include $path;",
        )
        .expect("write main php");
        std::fs::write(
            &child_path,
            "<?php $name = 'grandchild'; return include __DIR__ . '/' . $name . '.php';",
        )
        .expect("write child php");
        std::fs::write(
            &grandchild_path,
            "<?php return strpos(__FILE__, 'grandchild.php') !== false ? 42 : 0;",
        )
        .expect("write grandchild php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run nested php dynamic include");

        match vm.globals.get("$result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected nested $result global, got {other:?}"),
        }
    }

    #[test]
    fn php_dynamic_include_normalizes_alternative_control_syntax() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-alt-syntax-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let compat_path = base.join("compat.php");

        std::fs::write(
            &main_path,
            "<?php $result = require __DIR__ . '/compat.php'; $call_result = compat_polyfill();",
        )
        .expect("write main php");
        std::fs::write(
            &compat_path,
            "<?php\nif ( ! function_exists( 'compat_polyfill' ) ) :\n\t/**\n\t * Timing attack safe string comparison.\n\t *\n\t * @param string $known_string Expected string.\n\t * @param string $user_string  Actual, user supplied, string.\n\t * @return bool Whether strings are equal.\n\t */\n\tfunction compat_polyfill( $known_string = 'a', $user_string = 'a' ) {\n\t\t$known_string_length = strlen( $known_string );\n\n\t\tif ( strlen( $user_string ) !== $known_string_length ) {\n\t\t\treturn false;\n\t\t}\n\n\t\t$result = 0;\n\n\t\tfor ( $i = 0; $i < $known_string_length; $i++ ) {\n\t\t\t$result |= ord( $known_string[ $i ] ) ^ ord( $user_string[ $i ] );\n\t\t}\n\n\t\treturn 0 === $result ? 42 : 0;\n\t}\nendif;\n\nreturn compat_polyfill();\n",
        )
        .expect("write compat php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php with alternative syntax include");

        match vm.globals.get("$result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected include $result global, got {other:?}"),
        }

        match vm.globals.get("$call_result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected include $call_result global, got {other:?}"),
        }
    }

    #[test]
    fn php_dynamic_include_uses_entry_script_dir_for_nested_bare_relative_paths() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-entry-relative-{stamp}"));
        let classes_dir = base.join("classes");
        let views_dir = base.join("views");
        std::fs::create_dir_all(&classes_dir).expect("create classes dir");
        std::fs::create_dir_all(&views_dir).expect("create views dir");

        let main_path = base.join("main.php");
        let child_path = classes_dir.join("child.php");
        let partial_path = views_dir.join("partial.php");

        std::fs::write(
            &main_path,
            "<?php $path = 'classes/child.php'; $result = include $path;",
        )
        .expect("write main php");
        std::fs::write(
            &child_path,
            "<?php $view = 'views/partial.php'; return include $view;",
        )
        .expect("write child php");
        std::fs::write(&partial_path, "<?php return 42;")
            .expect("write partial php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php with entry-relative nested include");

        match vm.globals.get("$result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected nested entry-relative $result global, got {other:?}"),
        }
    }

    #[test]
    fn php_dynamic_include_renders_nested_view_templates_with_entry_relative_paths() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-view-render-{stamp}"));
        let views_projects_dir = base.join("views/projects");
        let views_templates_dir = base.join("views/templates");
        std::fs::create_dir_all(&views_projects_dir).expect("create projects view dir");
        std::fs::create_dir_all(&views_templates_dir).expect("create templates dir");

        let main_path = base.join("index.php");
        let list_path = views_projects_dir.join("list.php");
        let header_path = views_templates_dir.join("header.php");

        std::fs::write(
            &main_path,
            "<?php $result = include 'views/projects/list.php';",
        )
        .expect("write main php");
        std::fs::write(
            &list_path,
            "<?php $appName = 'Plan'; include 'views/templates/header.php'; ?>\n<main>Projects</main>\n",
        )
        .expect("write list php");
        std::fs::write(
            &header_path,
            "<header><?= htmlspecialchars($appName) ?></header>\n",
        )
        .expect("write header php");

        let rendered = Arc::new(Mutex::new(String::new()));
        let rendered_sink = rendered.clone();

        let mut vm = configured_vm();
        vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx, args| {
            let mut out = rendered_sink.lock().expect("lock rendered output");
            for arg in args {
                match arg {
                    Value::String(text) => out.push_str(text.as_ref()),
                    other => out.push_str(&format!("{}", other)),
                }
            }
            Value::Null
        }));

        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php nested view include");

        let rendered = rendered.lock().expect("lock rendered output").clone();
        assert!(rendered.contains("<header>Plan</header>"), "missing rendered header: {rendered}");
        assert!(rendered.contains("<main>Projects</main>"), "missing rendered list body: {rendered}");

    }
}