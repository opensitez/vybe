// vybe-test: csharp/common_patterns/out_var_declaration
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

if (int.TryParse("123", out var result)) {
    __Check((result).ToString(), "123");
}
