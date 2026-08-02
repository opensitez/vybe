// vybe-test: csharp/csharp_index_range_slice/array_range_one_before_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={2,4,6,8,10}; var slice=data[..^1]; __Check((slice[slice.Length-1]).ToString(), "8");
