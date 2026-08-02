// vybe-test: csharp/csharp_pattern_list/switch_expression_list_single_discard_arm_labels_single
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Label(int[] a)=>a switch{[_]=>"one",_=>"other"}; __Check((Label(new[]{9})).ToString(), "one");
