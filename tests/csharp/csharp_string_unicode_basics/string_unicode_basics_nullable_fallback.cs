// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
int? maybe = null; int fallback = maybe ?? 19; __Check((fallback == 19).ToString(), "True");
