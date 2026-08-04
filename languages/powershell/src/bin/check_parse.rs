use pest::Parser;
use vybe_language_powershell::{PowerShellParser, Rule};

fn main() {
    let src = std::env::args().nth(1).unwrap_or_else(|| "if ($a -ne $b) { }".to_string());
    println!("src: {}", src.replace('\n', "\\n"));

    let rules = [
        Rule::program,
        Rule::statement,
        Rule::command_stmt,
        Rule::pipeline_expr,
        Rule::command_segment,
        Rule::command_head,
        Rule::command_arg,
        Rule::command_token,
        Rule::quoted_string,
        Rule::if_stmt,
        Rule::condition_expr,
        Rule::condition_body,
        Rule::block,
    ];

    for rule in rules {
        match PowerShellParser::parse(rule, &src) {
            Ok(mut p) => {
                println!("{:?}: ok", rule);
                if let Some(root) = p.next() {
                    println!("  root: {}", root.as_str().replace('\n', "\\n"));
                    println!("  children: {}", root.clone().into_inner().count());
                }
            }
            Err(e) => println!("{:?}: err: {}", rule, e),
        }
    }
}
