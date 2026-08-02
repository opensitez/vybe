// vybe-test: csharp/modern_features/null_conditional_chain
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner { public string Value = "found"; }
class Outer { public Inner Child; }
var o = new Outer();
__Check((o.Child?.Value ?? "missing").ToString(), "missing");
o.Child = new Inner();
__Check((o.Child?.Value ?? "missing").ToString(), "found");
