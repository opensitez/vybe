// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_static_local_context
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtil {
    public static int Square(int n) {
        Log();
        return n * n;
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
__P((MathUtil.Square(4)).ToString());
__Check("Square\n16");
