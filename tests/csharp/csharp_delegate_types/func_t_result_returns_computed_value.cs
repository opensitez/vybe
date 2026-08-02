// vybe-test: csharp/csharp_delegate_types/func_t_result_returns_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> square = x => x * x;
__Check((square(4)).ToString(), "16");
