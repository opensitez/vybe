// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_capacity_unchanged_when_content_fits
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder(32); sb.Append("tiny"); __Check((sb.Capacity).ToString(), "32");
