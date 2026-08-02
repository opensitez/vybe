// vybe-test: csharp/csharp_caller_info_attributes/caller_file_path_and_member_name_together
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
        [System.Runtime.CompilerServices.CallerFilePath] string path = "") {
        __Check((member).ToString(), "Go");
        __Check((path.Length > 0).ToString(), "True");
    }
}
class App { public void Go() { Trace.Show(); } }
new App().Go();
