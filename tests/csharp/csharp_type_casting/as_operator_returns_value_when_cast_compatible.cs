// vybe-test: csharp/csharp_type_casting/as_operator_returns_value_when_cast_compatible
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = "hello"; string s = o as string; __Check((s).ToString(), "hello");
