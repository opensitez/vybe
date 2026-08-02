// vybe-test: csharp/csharp_pattern_list/is_list_not_pattern_accepts_different_shape
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=new[]{1}; __Check((data is not [1,2]).ToString(), "True");
