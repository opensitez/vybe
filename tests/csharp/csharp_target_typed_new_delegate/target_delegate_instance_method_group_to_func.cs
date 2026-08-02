// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_instance_method_group_to_func
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Scale { public int factor = 2; public int Apply(int n) => n * factor; }
System.Func<int, int> fn = new Scale().Apply;
__Check((fn(5)).ToString(), "10");
