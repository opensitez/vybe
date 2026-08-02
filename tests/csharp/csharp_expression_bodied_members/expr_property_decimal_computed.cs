// vybe-test: csharp/csharp_expression_bodied_members/expr_property_decimal_computed
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Price { public decimal unit = 2.5m; public decimal Triple => unit * 3m; }
__Check((new Price().Triple).ToString(), "7.5");
