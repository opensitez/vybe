// vybe-test: csharp/csharp_string_split_join/string_split_join_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
double seed = 21; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
