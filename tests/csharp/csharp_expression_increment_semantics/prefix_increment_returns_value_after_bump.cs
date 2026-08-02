// vybe-test: csharp/csharp_expression_increment_semantics/prefix_increment_returns_value_after_bump
// origin: languages/csharp/tests/csharp/test_csharp_expression_increment_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 4;
int read = ++n;
__Check((read).ToString(), "5");
__Check((n).ToString(), "5");
