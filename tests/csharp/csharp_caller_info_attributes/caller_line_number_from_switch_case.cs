// vybe-test: csharp/csharp_caller_info_attributes/caller_line_number_from_switch_case
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => __Check((line).ToString(), "5");
}
switch (1) {
    case 1: Trace.Show(); break;
}
