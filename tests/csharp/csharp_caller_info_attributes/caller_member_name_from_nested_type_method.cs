// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_nested_type_method
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    public class Inner { public void Work() { Trace.Show(); } }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Work");
}
new Outer.Inner().Work();
