// vybe-test: csharp/csharp_caller_info_attributes/caller_all_three_attributes_combined
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show(
        [System.Runtime.CompilerServices.CallerMemberName] string member = "",
        [System.Runtime.CompilerServices.CallerLineNumber] int line = 0,
        [System.Runtime.CompilerServices.CallerFilePath] string path = "") {
        __Check((member).ToString(), "Run");
        __Check((line).ToString(), "8");
        __Check((path.Length > 0).ToString(), "True");
    }
}
class App { public void Run() { Trace.Show(); } }
new App().Run();
