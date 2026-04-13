//! vybex — Universal compiler. Supports any language registered in languages/mod.rs.
//!
//! Usage: vybex <file>
//!
//! Language is detected from file extension (defined in each language's profile).

use std::path::Path;
use vybe_bytecode::VM;
use vybex::ast::*;

/// Resolve `import { x } from "./file.js"` by parsing the imported file
/// and prepending its body to the main module. Exported names become
/// globals that the main module can reference. Recursive (handles
/// transitive imports).
fn resolve_imports(module: &mut Module, lang: &vybex::languages::Language, base_dir: &Path) {
    let mut prepend: Vec<Statement> = Vec::new();
    for imp in &module.imports {
        let path_str = match &imp.kind {
            ImportKind::Named { path, .. } => path.clone(),
            ImportKind::Default { path, .. } => path.clone(),
            ImportKind::Simple { path, .. } => path.clone(),
            ImportKind::Wildcard { path, .. } => path.clone(),
        };
        // Resolve relative path
        let resolved = base_dir.join(&path_str);
        let source = match std::fs::read_to_string(&resolved) {
            Ok(s) => s,
            Err(e) => {
                eprintln!("Warning: cannot resolve import '{}': {}", path_str, e);
                continue;
            }
        };
        // Parse the imported module
        let mut imported = match (lang.parse)(&source) {
            Ok(m) => m,
            Err(e) => {
                eprintln!("Warning: parse error in '{}': {}", path_str, e);
                continue;
            }
        };
        // Recursively resolve nested imports
        let import_dir = resolved.parent().unwrap_or(base_dir);
        resolve_imports(&mut imported, lang, import_dir);
        // Prepend the imported module's body
        prepend.extend(imported.body);
    }
    // Insert imported code before the main body
    prepend.append(&mut module.body);
    module.body = prepend;
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        let exts = vybex::languages::supported_extensions();
        let ext_list: Vec<String> = exts.iter().map(|e| format!(".{}", e)).collect();
        eprintln!("Usage: vybex <file>\nSupported: {}", ext_list.join(", "));
        std::process::exit(1);
    }

    let path = Path::new(&args[1]);
    let source = match std::fs::read_to_string(path) {
        Ok(s) => s,
        Err(e) => { eprintln!("Cannot read {}: {}", path.display(), e); std::process::exit(1); }
    };

    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");

    // Find language by extension (reads [info].extensions from each profile)
    let lang = match vybex::languages::find_by_extension(ext) {
        Some(l) => l,
        None => {
            let exts = vybex::languages::supported_extensions();
            let ext_list: Vec<String> = exts.iter().map(|e| format!(".{}", e)).collect();
            eprintln!("Unknown file extension: .{}\nSupported: {}", ext, ext_list.join(", "));
            std::process::exit(1);
        }
    };

    // Parse source → common AST
    let mut module = match (lang.parse)(&source) {
        Ok(m) => m,
        Err(e) => { eprintln!("Parse error: {}", e); std::process::exit(1); }
    };

    // Resolve imports: parse each imported module and inline its body
    // before the main module's body. Named imports bind the exported
    // globals so the main module can reference them.
    let base_dir = path.parent().unwrap_or(Path::new("."));
    resolve_imports(&mut module, &lang, base_dir);

    // Load profile from embedded source
    let profile = match vybex::profile::parse_profile((lang.profile_source)()) {
        Ok(p) => p,
        Err(e) => { eprintln!("Profile error: {}", e); std::process::exit(1); }
    };

    // Compile AST → bytecode
    let chunks = match vybex::compiler::Compiler::with_profile(profile).compile(&module) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {}", e); std::process::exit(1); }
    };

    // Run on VM
    let mut vm = VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {}", e); std::process::exit(1); }
    }

    if gui.lock().unwrap().should_run {
        vybe_cli::runner::launch_vm_form(vm, gui, None);
    }
}
