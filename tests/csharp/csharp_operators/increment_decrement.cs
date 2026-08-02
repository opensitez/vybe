// vybe-test: csharp/csharp_operators/increment_decrement
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
__Check((x++).ToString(), "5");
__Check((x).ToString(), "6");
__Check((++x).ToString(), "7");
__Check((x--).ToString(), "7");
__Check((x).ToString(), "6");
