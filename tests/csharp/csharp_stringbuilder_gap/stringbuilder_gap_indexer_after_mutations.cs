// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_indexer_after_mutations
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("abc"); sb[1]='B'; sb.Append("d"); __Check((sb[0]).ToString(), "a"); __Check((sb.ToString()).ToString(), "aBcd");
