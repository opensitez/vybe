// vybe-test: csharp/csharp_text_stringbuilder/string_builder_append_line_adds_newline_separator
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder();
sb.AppendLine("line1").AppendLine("line2");
__Check((sb.ToString().Trim().Replace("\r\n","\n")).ToString(), "line1\nline2");
