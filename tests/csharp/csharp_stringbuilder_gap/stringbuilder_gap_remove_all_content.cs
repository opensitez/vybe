// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_remove_all_content
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("hello"); sb.Remove(0,5); __Check((sb.Length).ToString(), "0");
