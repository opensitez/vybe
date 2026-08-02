// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_static_method
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Ping");
    public static void Ping() { Show(); }
}
Trace.Ping();
