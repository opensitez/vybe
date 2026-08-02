// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_hashset_inferred_element_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.HashSet<int> set = new();
set.Add(1); set.Add(1); set.Add(2);
__Check((set.Count).ToString(), "2");
