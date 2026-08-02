// vybe-test: csharp/csharp_index_range_slice/range_both_indices_from_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={10,20,30,40,50}; var slice=data[^4..^1]; __Check((slice.Length).ToString(), "3"); __Check((slice[0]).ToString(), "20"); __Check((slice[2]).ToString(), "40");
