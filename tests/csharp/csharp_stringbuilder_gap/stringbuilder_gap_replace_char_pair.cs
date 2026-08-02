// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_replace_char_pair
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("x1x2"); sb.Replace('x','y'); __Check((sb.ToString()).ToString(), "y1y2");
