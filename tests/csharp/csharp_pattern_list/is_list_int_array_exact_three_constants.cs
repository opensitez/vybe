// vybe-test: csharp/csharp_pattern_list/is_list_int_array_exact_three_constants
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=new[]{2,4,6}; __Check((data is [2,4,6]).ToString(), "True");
