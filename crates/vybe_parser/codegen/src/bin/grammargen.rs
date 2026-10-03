//! Build-time CLI. Native Rust generation; no guest parser/runtime linkage.
use std::{env, fs, process::ExitCode};

fn main() -> ExitCode {
    let paths: Vec<_> = env::args_os().skip(1).collect();
    if paths.len() != 2 {
        eprintln!("usage: grammargen <input.grammar> <output.rs>");
        return ExitCode::from(2);
    }
    let source = match fs::read_to_string(&paths[0]) {
        Ok(source) => source,
        Err(error) => {
            eprintln!("{}: {error}", paths[0].to_string_lossy());
            return ExitCode::FAILURE;
        }
    };
    let generated = match vybe_parser_codegen::generate(&source) {
        Ok(generated) => generated,
        Err(errors) => {
            for error in errors {
                eprintln!("{}", error.render(&paths[0].to_string_lossy(), &source));
            }
            return ExitCode::FAILURE;
        }
    };
    // Preserve timestamps on unchanged output to avoid needless downstream
    // recompilation when a build script invokes the generator repeatedly.
    if fs::read_to_string(&paths[1]).ok().as_deref() == Some(&generated) {
        return ExitCode::SUCCESS;
    }
    match fs::write(&paths[1], generated) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{}: {error}", paths[1].to_string_lossy());
            ExitCode::FAILURE
        }
    }
}
