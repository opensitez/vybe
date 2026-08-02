// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_null_coalescing_param
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Safe { public string OrEmpty(string? s) => s ?? ""; }
__Check((new Safe().OrEmpty(null)).ToString(), ""); __Check((new Safe().OrEmpty("x")).ToString(), "x");
