// vybe-test: csharp/csharp_pattern_relational/is_relational_object_boxed_int_greater
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=9; __Check((o is int n and >5).ToString(), "True");
