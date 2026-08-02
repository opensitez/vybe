// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_replace_longer_with_shorter
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("a->b"); sb.Replace("->","-"); __Check((sb.ToString()).ToString(), "a-b");
