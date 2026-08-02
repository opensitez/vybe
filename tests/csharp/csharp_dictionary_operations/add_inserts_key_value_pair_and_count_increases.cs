// vybe-test: csharp/csharp_dictionary_operations/add_inserts_key_value_pair_and_count_increases
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>();
d.Add("a",1); d.Add("b",2);
__Check((d.Count).ToString(), "2");
