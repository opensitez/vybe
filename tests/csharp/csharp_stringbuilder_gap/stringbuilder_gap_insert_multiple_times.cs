// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_insert_multiple_times
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("a"); sb.Insert(1,"b").Insert(2,"c"); __Check((sb.ToString()).ToString(), "abc");
