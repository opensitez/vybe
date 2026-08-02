// vybe-test: csharp/csharp_delegate_types/func_stored_in_variable_and_passed_to_method
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Apply(System.Func<int,int> f, int v) => f(v);
__Check((Apply(x => x + 1, 9)).ToString(), "10");
