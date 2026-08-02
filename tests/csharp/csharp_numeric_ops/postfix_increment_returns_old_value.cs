// vybe-test: csharp/csharp_numeric_ops/postfix_increment_returns_old_value
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=5; int y=x++;
__Check((y).ToString(), "5"); __Check((x).ToString(), "6");
