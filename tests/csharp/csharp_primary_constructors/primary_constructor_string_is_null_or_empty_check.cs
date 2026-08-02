// vybe-test: csharp/csharp_primary_constructors/primary_constructor_string_is_null_or_empty_check
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Name(string? text) { public bool Missing => string.IsNullOrEmpty(text); }
__Check((new Name("").Missing).ToString(), "True");
