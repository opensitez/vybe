// vybe-test: csharp/csharp_new_features/target_typed_new_infers_list_type_from_variable
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> nums = new();
nums.Add(1); nums.Add(2);
__Check((nums.Count).ToString(), "2");
