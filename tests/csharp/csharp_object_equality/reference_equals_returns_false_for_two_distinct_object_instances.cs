// vybe-test: csharp/csharp_object_equality/reference_equals_returns_false_for_two_distinct_object_instances
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a = new object();
var b = new object();
__Check((object.ReferenceEquals(a, b)).ToString(), "False");
