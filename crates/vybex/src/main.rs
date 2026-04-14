//! vybex — Universal compiler.
//!
//! Usage: vybex <file|project>
//!
//! Supports single source files (detected by extension) and project files
//! (.vybe, .vbproj). Language is determined automatically.

use std::path::Path;
use vybe_bytecode::VM;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        let exts = vybex::projects::supported_extensions();
        let ext_list: Vec<String> = exts.iter().map(|e| format!(".{e}")).collect();
        eprintln!("Usage: vybex <file>\nSupported: {}", ext_list.join(", "));
        std::process::exit(1);
    }

    let path = Path::new(&args[1]);

    let bundle = match vybex::projects::load(path) {
        Ok(b) => b,
        Err(e) => { eprintln!("{e}"); std::process::exit(1); }
    };

    eprintln!("[vybex] Project '{}', sources={}, entry={:?}",
        bundle.name,
        bundle.sources.len(),
        bundle.entry_point);
    for s in &bundle.sources {
        eprintln!("  → {} ({} bytes)", s.path.display(), s.code.len());
    }

    let chunks = match bundle.compile() {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    if gui.lock().unwrap().should_run {
        vybe_host::gui_launch::launch_gui(vm, gui);
    }
}
