// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_case_sensitive_string_key_distinction
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["Mix"] = 2 }; __Check((map.ContainsKey("Mix")).ToString(), "True"); __Check((map.ContainsKey("mix")).ToString(), "False");
