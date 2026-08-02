// vybe-test: csharp/csharp_delegate_variance/action_object_to_string_contravariant_accepts_null
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> sink=v=>__Check((v==null).ToString(), "True"); System.Action<string> sinkString=sink; sinkString(null);
