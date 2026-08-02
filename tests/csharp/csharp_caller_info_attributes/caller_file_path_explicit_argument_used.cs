// vybe-test: csharp/csharp_caller_info_attributes/caller_file_path_explicit_argument_used
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerFilePath] string path = "") => __Check((path).ToString(), "/tmp/sample.cs");
}
Trace.Show("/tmp/sample.cs");
