mod editor;
pub mod form_designer_tab;
pub mod ide_text;
pub mod language;
mod lsp_client;
mod renderer;

use editor::Editor;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let sample = r#"// Sample Rust
fn main() {
    let x = 42;
    println!("x = {}", x);
}
"#;

    let lang = match crate::language::load_language("rust") {
        Some(l) => l,
        None => {
            eprintln!(
                "code_editor: language 'rust' not found. Expected JSON in one of:\n  - crates/code_editor/basic-languages/<name>/<name>.json\n  - basic-languages/<name>/<name>.json\n  - ../code_editor/basic-languages/<name>/<name>.json\nPlease ensure the 'basic-languages' folder is present relative to the binary working directory."
            );
            std::process::exit(1);
        }
    };
    let mut ed = Editor::from_text(sample, &lang);

    let is_gui =
        args.len() == 1 || args.contains(&"gui".to_string()) || args.contains(&"form".to_string());

    if is_gui {
        let open_form = args.contains(&"form".to_string());
        renderer::run_gui(ed, open_form);
        return;
    }

    // CLI fallback: print token summary and folds (run with `cli`)
    let tokens = ed.tokenize_all();
    println!("code_editor prototype (pure Rust) - token list:\n");
    for t in tokens.iter() {
        println!(
            "{:>5}-{:>5} {:<12} '{}'",
            t.start,
            t.end,
            format!("{:?}", t.kind),
            ed.slice(t.start, t.end).replace('\n', "\\n")
        );
    }
    println!("\nFolds: {:?}", ed.folds());
    // Demo edit then retokenize
    ed.insert_str(10, "    // inserted comment\n", &lang);
    println!(
        "After edit: tokens={}, folds={}",
        ed.tokenize_all().len(),
        ed.folds().len()
    );
}
