// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_appendformat_two_placeholders
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder(); sb.AppendFormat("{0}-{1}",1,2); __Check((sb.ToString()).ToString(), "1-2");
