// vybe-test: csharp/csharp_delegate_types/method_group_assigned_to_func_without_lambda_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string,int> len = s => s.Length;
__Check((len("test")).ToString(), "4");
