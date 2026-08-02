// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Run");
}
class App { public void Run() { Trace.Show(); } }
new App().Run();
