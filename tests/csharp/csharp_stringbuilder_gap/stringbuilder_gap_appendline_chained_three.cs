// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_appendline_chained_three
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder(); sb.AppendLine("a").AppendLine("b").AppendLine("c"); __Check((sb.ToString().Replace("\r\n","\n").Trim().Split('\n').Length).ToString(), "3");
