// vybe-test: csharp/csharp_new_features/nameof_returns_string_name_of_variable
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int myCounter = 0;
__Check((nameof(myCounter)).ToString(), "myCounter");
