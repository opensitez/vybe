// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_operator_method
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Num {
    public int V;
    public static Num operator +(Num a, Num b) {
        Log();
        return new Num { V = a.V + b.V };
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "op_Addition");
}
__Check(((new Num { V = 1 } + new Num { V = 2 }).V).ToString(), "3");
