// vybe-test: csharp/modern_features/is_not_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object obj = "test";
if (obj is not null) {
    __Check(("not null").ToString(), "not null");
}
if (obj is not int) {
    __Check(("not int").ToString(), "not int");
}
