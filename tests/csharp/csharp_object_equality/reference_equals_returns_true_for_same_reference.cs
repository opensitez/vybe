// vybe-test: csharp/csharp_object_equality/reference_equals_returns_true_for_same_reference
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a = new object();
var b = a;
__Check((object.ReferenceEquals(a, b)).ToString(), "True");
