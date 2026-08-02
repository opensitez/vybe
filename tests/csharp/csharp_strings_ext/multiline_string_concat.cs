// vybe-test: csharp/csharp_strings_ext/multiline_string_concat
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "Hello" +
    " " +
    "World";
__Check((result).ToString(), "Hello World");
