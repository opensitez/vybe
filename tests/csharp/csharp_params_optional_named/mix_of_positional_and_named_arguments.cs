// vybe-test: csharp/csharp_params_optional_named/mix_of_positional_and_named_arguments
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Sub(int x, int y) => x-y;
__Check((Sub(10, y:3)).ToString(), "7");
