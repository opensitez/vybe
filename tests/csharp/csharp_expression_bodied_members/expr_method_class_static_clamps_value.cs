// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_static_clamps_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class ClampUtil { public static int Clamp(int v, int lo, int hi) => v < lo ? lo : v > hi ? hi : v; }
__Check((ClampUtil.Clamp(15, 0, 10)).ToString(), "10");
