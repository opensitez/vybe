// vybe-test: csharp/csharp_string_split_join/string_split_join_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
