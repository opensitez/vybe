// vybe-test: csharp/csharp_delegate_variance/action_contravariant_method_group
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static void Print(object o)=>__Check((o).ToString(), "group"); System.Action<object> wide=Print; System.Action<string> narrow=wide; narrow("group");
