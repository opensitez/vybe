// vybe-test: csharp/csharp_string_split_join/string_split_join_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
string feature = "string_split_join"; __Check((feature.Length > 0).ToString(), "True");
