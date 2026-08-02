// vybe-test: csharp/csharp_pattern_list/switch_expression_list_returns_string_from_var_slots
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string PairLabel(int[] a)=>a switch{[var x,var y]=>$"{x}-{y}",_=>"?"}; __Check((PairLabel(new[]{7,8})).ToString(), "7-8");
