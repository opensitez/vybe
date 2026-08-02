// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_dictionary_with_object_initializer_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Dictionary<string, int> map = new() { ["a"] = 1, ["b"] = 2 };
__Check((map["b"]).ToString(), "2");
