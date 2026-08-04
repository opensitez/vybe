// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_operator_method
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

class Num {
    public int V;
    public static Num operator +(Num a, Num b) {
        Log();
        return new Num { V = a.V + b.V };
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
__P(((new Num { V = 1 } + new Num { V = 2 }).V).ToString());
__Check("op_Addition\n3");
