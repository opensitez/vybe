// vybe-test: csharp/csharp_collections_generic/dictionary_contains_key_vs_contains_value
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=new System.Collections.Generic.Dictionary<string,int>{{"a",1}};
__Check((d.ContainsKey("a")).ToString(), "True");
__Check((d.ContainsValue(1)).ToString(), "True");
__Check((d.ContainsKey("b")).ToString(), "False");
