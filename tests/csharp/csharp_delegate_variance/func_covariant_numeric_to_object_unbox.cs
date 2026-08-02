// vybe-test: csharp/csharp_delegate_variance/func_covariant_numeric_to_object_unbox
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int> f=()=>123; System.Func<object> g=f; __Check(((int)g()).ToString(), "123");
