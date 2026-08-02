// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_appendline_then_replace_newline
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder(); sb.AppendLine("row"); __Check((sb.ToString().Contains("\n")||sb.ToString().Contains("\r")).ToString(), "True");
