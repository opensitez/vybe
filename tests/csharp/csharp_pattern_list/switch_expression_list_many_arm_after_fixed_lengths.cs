// vybe-test: csharp/csharp_pattern_list/switch_expression_list_many_arm_after_fixed_lengths
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Size(int[] a)=>a switch{[]=>"0",[_]=>"1",[_,_]=>"2",_=>"many"}; __Check((Size(new[]{1,2,3})).ToString(), "many");
