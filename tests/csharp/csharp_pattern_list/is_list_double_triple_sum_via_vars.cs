// vybe-test: csharp/csharp_pattern_list/is_list_double_triple_sum_via_vars
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double[] vals=new[]{1.5,2.0,2.5}; if(vals is [var a,var b,var c]) __Check((a+b+c).ToString(), "6");
