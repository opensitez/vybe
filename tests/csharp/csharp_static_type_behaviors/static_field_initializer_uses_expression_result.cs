// vybe-test: csharp/csharp_static_type_behaviors/static_field_initializer_uses_expression_result
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Limits {
    public static int Max = 8 * 8;
}
__Check((Limits.Max).ToString(), "64");
