// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_appendline_no_argument
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("a"); sb.AppendLine(); __Check((sb.Length).ToString(), "3");
