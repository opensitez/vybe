use std::env;

use vybe_language_powershell::{PowerShellParser, Rule};
use pest::Parser;

fn main() {
    let src = env::args()
        .nth(1)
        .unwrap_or_else(|| "if ($true) {}".to_string());

    println!("--- program ---");
    match PowerShellParser::parse(Rule::program, &src) {
        Ok(mut p) => {
            println!("program ok");
            if let Some(root) = p.next() {
                for child in root.into_inner() {
                    println!("program child: {:?}", child.as_rule());
                }
            }
        }
        Err(e) => {
            println!("program err: {e}");
        }
    }

    println!("--- if_stmt ---");
    match PowerShellParser::parse(Rule::if_stmt, &src) {
        Ok(mut p) => {
            println!("if_stmt ok: {}", p.next().map(|p| p.as_str().replace('\n', "\\n")).unwrap_or_default());
            if let Some(root) = p.next() {
                println!("unexpected extra {} pairs", root.into_inner().count());
            }
        }
        Err(e) => println!("if_stmt err: {e}"),
    }

    println!("--- command_head ---");
    match PowerShellParser::parse(Rule::command_head, &src) {
        Ok(mut p) => println!("command_head ok: {}", p.next().unwrap().as_str()),
        Err(e) => println!("command_head err: {e}"),
    }

    println!("--- condition_expr ---");
    match PowerShellParser::parse(Rule::condition_expr, &src) {
        Ok(mut p) => println!(
            "condition_expr ok: {}",
            p.next().unwrap().as_str().replace('\n', "\\n")
        ),
        Err(e) => println!("condition_expr err: {e}"),
    }

    println!("--- condition_body ---");
    match PowerShellParser::parse(Rule::condition_body, &src) {
        Ok(mut p) => println!(
            "condition_body ok: {}",
            p.next().unwrap().as_str().replace('\n', "\\n")
        ),
        Err(e) => println!("condition_body err: {e}"),
    }
}
