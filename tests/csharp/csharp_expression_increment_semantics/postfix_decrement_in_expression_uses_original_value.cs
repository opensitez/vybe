// vybe-test: csharp/csharp_expression_increment_semantics/postfix_decrement_in_expression_uses_original_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_increment_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 3;
int total = n-- + n;
__Check((total).ToString(), "5");
__Check((n).ToString(), "2");
