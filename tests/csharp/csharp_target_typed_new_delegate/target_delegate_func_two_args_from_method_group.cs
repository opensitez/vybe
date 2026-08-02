// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_two_args_from_method_group
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int Add(int a, int b) => a + b;
System.Func<int, int, int> add = Add;
__Check((add(3, 4)).ToString(), "7");
