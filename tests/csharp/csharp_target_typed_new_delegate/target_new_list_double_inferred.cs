// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_double_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<double> nums = new();
nums.Add(1.5); nums.Add(2.5);
__Check((nums[0] + nums[1]).ToString(), "4");
