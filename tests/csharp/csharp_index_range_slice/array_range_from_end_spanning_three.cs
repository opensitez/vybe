// vybe-test: csharp/csharp_index_range_slice/array_range_from_end_spanning_three
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4,5}; var slice=data[^5..^2]; __Check((slice.Length).ToString(), "3"); __Check((slice[2]).ToString(), "3");
