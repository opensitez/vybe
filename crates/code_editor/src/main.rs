mod editor;
mod renderer;
pub mod language;

use editor::Editor;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let sample = r#"// Sample Rust
fn main() {
    let x = 42;
    println!("x = {}", x);
}
"#;

    let mut ed = Editor::from_text(sample);

    // Default to GUI. Pass `cli` as the first argument to force CLI output.
    if args.len() == 1 || (args.len() > 1 && args[1] == "gui") {
        renderer::run_gui(ed);
        return;
    }

    // CLI fallback: print token summary and folds (run with `cli`)
    let tokens = ed.tokenize_all();
    println!("code_editor prototype (pure Rust) - token list:\n");
    for t in tokens.iter() {
        println!("{:>5}-{:>5} {:<12} '{}'",
            t.start, t.end, format!("{:?}", t.kind), ed.slice(t.start, t.end).replace('\n', "\\n")
        );
    }
    println!("\nFolds: {:?}", ed.folds());
    // Demo edit then retokenize
    ed.insert_str(10, "    // inserted comment\n");
    println!("After edit: tokens={}, folds={}", ed.tokenize_all().len(), ed.folds().len());
}
