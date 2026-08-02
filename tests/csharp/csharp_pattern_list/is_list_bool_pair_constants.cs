// vybe-test: csharp/csharp_pattern_list/is_list_bool_pair_constants
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool[] flags=new[]{true,false}; __Check((flags is [true,false]).ToString(), "True");
