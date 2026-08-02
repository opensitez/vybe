// vybe-test: csharp/csharp_pattern_list/is_list_triple_constant_pattern_matches_literal_head
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{5,6,7}; __Check((data is [5,6,7]).ToString(), "True");
