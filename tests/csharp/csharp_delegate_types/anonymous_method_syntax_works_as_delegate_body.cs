// vybe-test: csharp/csharp_delegate_types/anonymous_method_syntax_works_as_delegate_body
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> triple = delegate(int n) { return n * 3; };
__Check((triple(3)).ToString(), "9");
