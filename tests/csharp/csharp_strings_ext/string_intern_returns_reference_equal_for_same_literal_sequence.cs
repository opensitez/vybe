// vybe-test: csharp/csharp_strings_ext/string_intern_returns_reference_equal_for_same_literal_sequence
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a = string.Intern("shared");
string b = string.Intern("shared");
__Check((object.ReferenceEquals(a, b)).ToString(), "True");
