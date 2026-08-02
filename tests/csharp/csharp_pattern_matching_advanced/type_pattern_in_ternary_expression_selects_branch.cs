// vybe-test: csharp/csharp_pattern_matching_advanced/type_pattern_in_ternary_expression_selects_branch
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = "cs"; __Check((item is string ? "text" : "other").ToString(), "text");
