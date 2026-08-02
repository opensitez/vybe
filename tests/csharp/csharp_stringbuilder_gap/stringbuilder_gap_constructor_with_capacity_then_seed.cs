// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_constructor_with_capacity_then_seed
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder(16); sb.Append("seed"); sb.Append("+"); __Check((sb.ToString()).ToString(), "seed+");
