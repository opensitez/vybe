// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
string feature = "char_predicate_apis"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
