// vybe-test: csharp/csharp_pattern_list/switch_expression_list_default_after_length_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Bucket(int[] a)=>a switch{[]=>"e",[_]=>"s",_=>"m"}; __Check((Bucket(new[]{1,2})).ToString(), "m");
