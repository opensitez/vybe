// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_remove_zero_count
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("keep"); sb.Remove(2,0); __Check((sb.ToString()).ToString(), "keep");
