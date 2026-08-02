// vybe-test: csharp/csharp_string_split_join/string_split_join_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
string feature = "string_split_join"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
