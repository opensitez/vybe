// vybe-test: csharp/csharp_index_range_slice/index_from_end_one_matches_length_minus_one
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={3,6,9}; __Check((data[data.Length-1]).ToString(), "9"); __Check((data[^1]).ToString(), "9");
