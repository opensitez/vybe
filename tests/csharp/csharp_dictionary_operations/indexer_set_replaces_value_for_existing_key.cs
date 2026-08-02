// vybe-test: csharp/csharp_dictionary_operations/indexer_set_replaces_value_for_existing_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>();
d["x"] = 1; d["x"] = 9;
__Check((d["x"]).ToString(), "9");
