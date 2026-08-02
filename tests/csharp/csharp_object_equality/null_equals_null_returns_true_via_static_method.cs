// vybe-test: csharp/csharp_object_equality/null_equals_null_returns_true_via_static_method
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((object.Equals(null, null)).ToString(), "True");
