// vybe-test: csharp/csharp_method_group_delegates/method_group_converts_to_func_without_explicit_lambda_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_method_group_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int Double(int n) => n * 2;
System.Func<int, int> fn = Double;
__Check((fn(6)).ToString(), "12");
