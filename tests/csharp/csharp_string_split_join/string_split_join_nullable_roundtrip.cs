// vybe-test: csharp/csharp_string_split_join/string_split_join_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
int? maybe = 21; __Check((maybe.HasValue && maybe.Value == 21).ToString(), "True");
