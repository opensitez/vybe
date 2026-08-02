// vybe-test: csharp/csharp_caller_info_attributes/caller_line_number_zero_when_explicit_zero
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = -1) => __Check((line).ToString(), "0");
}
Trace.Show(0);
