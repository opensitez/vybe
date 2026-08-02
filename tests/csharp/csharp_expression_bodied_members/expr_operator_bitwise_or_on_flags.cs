// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_bitwise_or_on_flags
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Bits { public int V; public static Bits operator |(Bits a, Bits b) => new Bits { V = a.V | b.V }; }
__Check(((new Bits { V = 1 } | new Bits { V = 2 }).V).ToString(), "3");
