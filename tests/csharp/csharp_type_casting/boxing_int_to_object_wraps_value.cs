// vybe-test: csharp/csharp_type_casting/boxing_int_to_object_wraps_value
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object boxed = 42; __Check((boxed).ToString(), "42");
