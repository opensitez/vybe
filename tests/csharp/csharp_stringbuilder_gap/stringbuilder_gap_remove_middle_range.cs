// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_remove_middle_range
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("abcde"); sb.Remove(1,3); __Check((sb.ToString()).ToString(), "ae");
