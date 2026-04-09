//! vybex — Universal compiler. Supports any language registered in languages/mod.rs.
//!
//! Usage: vybex <file>
//!
//! Language is detected from file extension (defined in each language's profile).

use std::path::Path;
use vybe_bytecode::VM;

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
    let module = match (lang.parse)(&source) {
        Ok(m) => m,
        Err(e) => { eprintln!("Parse error: {}", e); std::process::exit(1); }
    };

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
