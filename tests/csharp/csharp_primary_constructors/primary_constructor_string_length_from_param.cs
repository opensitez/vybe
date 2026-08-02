// vybe-test: csharp/csharp_primary_constructors/primary_constructor_string_length_from_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Label(string text) { public int Length => text.Length; }
__Check((new Label("abcd").Length).ToString(), "4");
