// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
int seed = 23; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
