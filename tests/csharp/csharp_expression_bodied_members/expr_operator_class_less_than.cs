// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_class_less_than
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Score { public int V; public static bool operator <(Score a, Score b) => a.V < b.V; public static bool operator >(Score a, Score b) => a.V > b.V; }
__Check((new Score { V = 1 } < new Score { V = 2 }).ToString(), "True"); __Check((new Score { V = 5 } > new Score { V = 3 }).ToString(), "True");
