// vybe-test: csharp/csharp_string_builder/append_line_adds_content_followed_by_newline
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder();
sb.AppendLine("line1");
__Check((sb.Length > 5).ToString(), "True");
