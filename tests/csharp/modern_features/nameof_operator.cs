// vybe-test: csharp/modern_features/nameof_operator
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int myVariable = 10;
__Check((nameof(myVariable)).ToString(), "myVariable");
__Check((nameof(Console)).ToString(), "Console");
