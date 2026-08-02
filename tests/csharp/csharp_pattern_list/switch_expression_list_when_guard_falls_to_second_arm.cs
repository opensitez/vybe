// vybe-test: csharp/csharp_pattern_list/switch_expression_list_when_guard_falls_to_second_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Rank(int[] a)=>a switch{[var x,var y] when x>y=>"desc",[var x,var y]=>"asc",_=>"other"}; __Check((Rank(new[]{2,5})).ToString(), "asc");
