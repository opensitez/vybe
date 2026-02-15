#[cfg(test)]
mod tests {
    use vybe_parser::parse_program;
    use vybe_runtime::Interpreter;
    use vybe_runtime::Value;

    #[test]
    fn test_xml_literal() {
        let code = r#"
            Module Test
                Sub Main()
                    Dim name = "World"
                    Dim x = <root>
                                <child id="1">Hello <%= name %></child>
                            </root>
                    
                    Console.WriteLine(x.ToString())
                End Sub
            End Module
        "#;
        let program = parse_program(code).unwrap();
        let mut interp = Interpreter::new();
        interp.run(&program).unwrap();
        // Verify output via side_effects if needed, or just ensure it runs without error
    }
}
