// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_insert_empty_string_noop
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("same"); sb.Insert(2,""); __Check((sb.ToString()).ToString(), "same");
