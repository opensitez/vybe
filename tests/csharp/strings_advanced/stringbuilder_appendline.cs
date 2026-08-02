// vybe-test: csharp/strings_advanced/stringbuilder_appendline
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

var sb = new System.Text.StringBuilder();
sb.AppendLine("line1");
sb.AppendLine("line2");
Console.Write(sb.ToString());
