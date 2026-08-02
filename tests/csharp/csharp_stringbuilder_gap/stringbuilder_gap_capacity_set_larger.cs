// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_capacity_set_larger
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("hi"); sb.Capacity=64; __Check((sb.Capacity>=64).ToString(), "True");
