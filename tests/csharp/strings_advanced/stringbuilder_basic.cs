// vybe-test: csharp/strings_advanced/stringbuilder_basic
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder();
sb.Append("Hello");
sb.Append(" ");
sb.Append("World");
__Check((sb.ToString()).ToString(), "Hello World");
