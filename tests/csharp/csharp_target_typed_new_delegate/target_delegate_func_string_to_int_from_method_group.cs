// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_string_to_int_from_method_group
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int Len(string s) => s.Length;
System.Func<string, int> measure = Len;
__Check((measure("hello")).ToString(), "5");
