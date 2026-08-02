// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
string feature = "char_predicate_apis:23"; __Check((feature.Length >= 1).ToString(), "True");
