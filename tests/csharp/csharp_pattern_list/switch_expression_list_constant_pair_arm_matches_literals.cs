// vybe-test: csharp/csharp_pattern_list/switch_expression_list_constant_pair_arm_matches_literals
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Code(int[] a)=>a switch{[1,2]=>"twelve",_=>"other"}; __Check((Code(new[]{1,2})).ToString(), "twelve");
