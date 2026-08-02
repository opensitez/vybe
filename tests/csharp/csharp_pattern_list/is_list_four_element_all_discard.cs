// vybe-test: csharp/csharp_pattern_list/is_list_four_element_all_discard
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=new[]{1,2,3,4}; __Check((data is [_,_,_,_]).ToString(), "True");
