// vybe-test: csharp/csharp_pattern_list/switch_expression_list_var_pair_arm_returns_sum
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int SumPair(int[] a)=>a switch{[var x,var y]=>x+y,_=>0}; __Check((SumPair(new[]{10,20})).ToString(), "30");
