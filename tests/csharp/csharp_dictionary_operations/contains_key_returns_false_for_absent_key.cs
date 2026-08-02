// vybe-test: csharp/csharp_dictionary_operations/contains_key_returns_false_for_absent_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",1}};
__Check((d.ContainsKey("z")).ToString(), "False");
