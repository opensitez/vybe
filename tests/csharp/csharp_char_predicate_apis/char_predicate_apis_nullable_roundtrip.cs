// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
int? maybe = 23; __Check((maybe.HasValue && maybe.Value == 23).ToString(), "True");
