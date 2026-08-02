// vybe-test: csharp/csharp_caller_info_attributes/caller_line_number_two_calls_different_lines
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
Trace.Show();
Trace.Show();
