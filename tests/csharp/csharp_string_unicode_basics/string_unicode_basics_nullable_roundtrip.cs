// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
int? maybe = 19; __Check((maybe.HasValue && maybe.Value == 19).ToString(), "True");
