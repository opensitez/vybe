// vybe-test: csharp/csharp_delegate_variance/func_covariant_two_independent_assignments
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> a=()=>"one"; System.Func<string> b=()=>"two"; System.Func<object> ga=a; System.Func<object> gb=b; __Check((ga()).ToString(), "one"); __Check((gb()).ToString(), "two");
