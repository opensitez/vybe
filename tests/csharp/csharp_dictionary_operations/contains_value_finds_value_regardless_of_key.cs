// vybe-test: csharp/csharp_dictionary_operations/contains_value_finds_value_regardless_of_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",42}};
__Check((d.ContainsValue(42)).ToString(), "True");
