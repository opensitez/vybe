// vybe-test: csharp/csharp_type_casting/unboxing_casts_object_back_to_int
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object boxed = 42; int n = (int)boxed; __Check((n).ToString(), "42");
