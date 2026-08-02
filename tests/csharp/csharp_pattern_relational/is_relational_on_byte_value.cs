// vybe-test: csharp/csharp_pattern_relational/is_relational_on_byte_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte b=200; __Check((b is >100).ToString(), "True");
