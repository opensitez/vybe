use std::env;

use vybe_language_powershell::parse;

fn main() {
    let src = env::args()
        .nth(1)
        .unwrap_or_else(|| "if ($true) {\n}\n".to_string());

    match parse(&src) {
        Ok(ast) => {
            println!("ok: {}", ast.body.len());
            for (idx, stmt) in ast.body.iter().enumerate() {
                println!("{idx}: {:?}", stmt.kind);
            }
        }
        Err(err) => {
            println!("err: {err}");
        }
    }
}
