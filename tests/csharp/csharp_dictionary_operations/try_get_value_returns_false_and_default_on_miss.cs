// vybe-test: csharp/csharp_dictionary_operations/try_get_value_returns_false_and_default_on_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>();
__Check((d.TryGetValue("nope", out int v)).ToString(), "False");
__Check((v).ToString(), "0");
