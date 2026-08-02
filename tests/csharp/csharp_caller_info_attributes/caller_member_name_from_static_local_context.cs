// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_static_local_context
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtil {
    public static int Square(int n) {
        Log();
        return n * n;
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Square");
}
__Check((MathUtil.Square(4)).ToString(), "16");
