// vybe-test: csharp/csharp_delegate_variance/func_covariant_chain_two_hops_to_object
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> inner=()=>"chain"; System.Func<object> mid=inner; System.Func<object> outer=mid; __Check((outer()).ToString(), "chain");
