// vybe-test: csharp/strings_advanced/stringbuilder_insert_replace
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("Hello World");
sb.Replace("World", "There");
__Check((sb.ToString()).ToString(), "Hello There");
sb.Insert(5, " Beautiful");
__Check((sb.ToString()).ToString(), "Hello Beautiful There");
