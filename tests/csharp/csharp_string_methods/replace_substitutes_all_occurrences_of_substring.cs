// vybe-test: csharp/csharp_string_methods/replace_substitutes_all_occurrences_of_substring
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("aabbaa".Replace("aa","X")).ToString(), "XbbX");
