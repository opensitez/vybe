// vybe-test: csharp/csharp_caller_info_attributes/caller_line_number_with_other_optional_params
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show(string tag, [System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => __Check((tag + ":" + line).ToString(), "mark:4");
}
Trace.Show("mark");
