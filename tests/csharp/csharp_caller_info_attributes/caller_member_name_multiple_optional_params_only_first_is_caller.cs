// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_multiple_optional_params_only_first_is_caller
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show(string prefix = "p", [System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((prefix + member).ToString(), "pRun");
}
class App { public void Run() { Trace.Show(); } }
new App().Run();
