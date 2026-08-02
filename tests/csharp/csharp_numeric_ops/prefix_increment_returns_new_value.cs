// vybe-test: csharp/csharp_numeric_ops/prefix_increment_returns_new_value
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=5; int y=++x;
__Check((y).ToString(), "6"); __Check((x).ToString(), "6");
