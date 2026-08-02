// vybe-test: csharp/csharp_operators/as_operator_returns_null_for_incompatible_reference_cast
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object value = 1;
var text = value as string;
__Check((text == null).ToString(), "True");
