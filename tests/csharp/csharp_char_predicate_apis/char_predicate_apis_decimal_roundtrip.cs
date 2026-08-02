// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
