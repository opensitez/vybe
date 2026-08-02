// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
double seed = 23; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
