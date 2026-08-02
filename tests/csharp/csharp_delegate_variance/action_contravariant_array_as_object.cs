// vybe-test: csharp/csharp_delegate_variance/action_contravariant_array_as_object
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> log=v=>__Check((v is int[]).ToString(), "True"); System.Action<int[]> logArr=log; logArr(new int[]{1});
