#[cfg(test)]
mod tests {
    use vybe_parser_basic::parse_program;
    use vybe_runtime::Interpreter;
    use vybe_runtime::Value;

    #[test]
    fn test_task_run() {
        let code = r#"
            Imports System.Threading.Tasks
            
            Module Test
                Sub Main()
                    Dim t = Task.Run(Sub() 
                                        Console.WriteLine("Task Running")
                                     End Sub)
                    t.Wait()
                    Console.WriteLine("Task Completed")
                End Sub
            End Module
        "#;
        let program = parse_program(code).unwrap();
        let mut interp = Interpreter::new();
        interp.run(&program).unwrap();
        // Verify execution order or completion
    }
}
