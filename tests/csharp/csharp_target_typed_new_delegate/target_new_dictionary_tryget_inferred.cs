// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_dictionary_tryget_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Dictionary<string, int> map = new();
map["x"] = 5;
__Check((map.TryGetValue("x", out int v)).ToString(), "True"); __Check((v).ToString(), "5");
