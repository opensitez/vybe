// vybe-test: csharp/csharp_expression_bodied_members/expr_method_long_arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Wide { public long Add(long a, long b) => a + b; }
__Check((new Wide().Add(9000000000L, 1L)).ToString(), "9000000001");
