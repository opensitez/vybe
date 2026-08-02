// vybe-test: csharp/csharp_string_builder/append_builds_string_incrementally
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder();
sb.Append("Hello"); sb.Append(" World");
__Check((sb.ToString()).ToString(), "Hello World");
