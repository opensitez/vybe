// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_length_after_remove
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("abcdef"); sb.Remove(2,2); __Check((sb.Length).ToString(), "4");
