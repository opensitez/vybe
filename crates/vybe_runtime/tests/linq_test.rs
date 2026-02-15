#[cfg(test)]
mod tests {
    use vybe_parser::parse_program;
    use vybe_runtime::Interpreter;
    use vybe_runtime::Value;

    fn run_code(code: &str) -> Value {
        let program = parse_program(code).unwrap();
        let mut interpreter = Interpreter::new();
        interpreter.run(&program).unwrap();
        Value::Nothing // run returns (), so return Nothing or capture output
    }

    #[test]
    fn test_linq_simple_select() {
        let code = r#"
            Module Test
                Sub Main()
                    Dim nums() = {1, 2, 3, 4, 5}
                    Dim q = From n In nums Where n > 2 Select n
                    
                    ' Verify result is array [3, 4, 5]
                    Dim count = q.Count
                    Console.WriteLine(count)
                    Console.WriteLine(q(0))
                End Sub
            End Module
        "#;
        run_code(code);
    }
    
    // Better: Helper to run script-like snippets if supported, or just use full program structure.
}
