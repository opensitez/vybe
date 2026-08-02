// vybe-test: csharp/csharp_pattern_list/switch_expression_list_empty_arm_labels_empty
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Label(int[] a)=>a switch{[]=>"empty",_=>"other"}; __Check((Label(new int[]{})).ToString(), "empty");
