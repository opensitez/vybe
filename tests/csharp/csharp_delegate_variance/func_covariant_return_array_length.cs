// vybe-test: csharp/csharp_delegate_variance/func_covariant_return_array_length
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int[]> f=()=>new int[]{1,2,3}; System.Func<object> g=f; __Check((((int[])g()).Length).ToString(), "3");
