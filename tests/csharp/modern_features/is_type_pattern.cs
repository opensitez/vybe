// vybe-test: csharp/modern_features/is_type_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object obj = "hello";
if (obj is string s) {
    __Check(("string: " + s.ToUpper()).ToString(), "string: HELLO");
}
