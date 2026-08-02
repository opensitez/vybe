// vybe-test: csharp/csharp_delegate_variance/action_contravariant_string_is_object_check
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> probe=v=>__Check((v is string).ToString(), "True"); System.Action<string> probeString=probe; probeString("ok");
