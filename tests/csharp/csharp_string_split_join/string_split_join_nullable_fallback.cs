// vybe-test: csharp/csharp_string_split_join/string_split_join_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
int? maybe = null; int fallback = maybe ?? 21; __Check((fallback == 21).ToString(), "True");
