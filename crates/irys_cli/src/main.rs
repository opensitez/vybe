use std::env;
use std::path::PathBuf;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: irys <filename.vb|filename.vbp|filename.vbproj> [args...]");
        std::process::exit(1);
    }

    let file_path = PathBuf::from(&args[1]);
    if !file_path.exists() {
        eprintln!("Error: file not found: {}", file_path.display());
        std::process::exit(1);
    }

    let extra_args: Vec<String> = args[2..].to_vec();
    irys_ui::run(&file_path, &extra_args);
}
