// vybe-test: csharp/csharp_dictionary_operations/remove_deletes_key_and_reduces_count
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",1},{"b",2}};
d.Remove("a");
__Check((d.Count).ToString(), "1");
