// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
int seed = 19; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
