// vybe-test: csharp/csharp_dictionary_operations/try_get_value_returns_true_and_out_value_on_hit
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>{{"k",5}};
__Check((d.TryGetValue("k", out int v)).ToString(), "True");
__Check((v).ToString(), "5");
