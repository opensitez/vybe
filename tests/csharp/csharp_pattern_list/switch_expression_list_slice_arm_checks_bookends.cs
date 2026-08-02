// vybe-test: csharp/csharp_pattern_list/switch_expression_list_slice_arm_checks_bookends
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Edge(int[] a)=>a switch{[1,..,9]=>"book",_=>"plain"}; __Check((Edge(new[]{1,5,9})).ToString(), "book");
