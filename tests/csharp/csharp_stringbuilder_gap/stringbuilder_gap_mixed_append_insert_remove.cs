// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_mixed_append_insert_remove
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("start"); sb.Append("_end"); sb.Insert(5,"-mid-"); sb.Remove(0,6); __Check((sb.ToString()).ToString(), "mid-_end");
