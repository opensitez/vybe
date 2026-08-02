// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_clear_then_append
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("old"); sb.Clear(); sb.Append("new"); __Check((sb.ToString()).ToString(), "new");
