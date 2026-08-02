// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_hashset_string_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.HashSet<string> tags = new();
tags.Add("a"); tags.Add("b");
__Check((tags.Contains("a")).ToString(), "True"); __Check((tags.Contains("c")).ToString(), "False");
