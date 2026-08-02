// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class App {
    public void Run() {
        Local();
        void Local() {
            Trace.Show();
        }
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Local");
}
new App().Run();
