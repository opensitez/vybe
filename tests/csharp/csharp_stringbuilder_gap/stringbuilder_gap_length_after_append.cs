// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_length_after_append
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder(); sb.Append("four"); __Check((sb.Length).ToString(), "4");
