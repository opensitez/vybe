// vybe-test: csharp/csharp_index_range_slice/range_end_from_end_start_open
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4}; System.Range r=..^1; var slice=data[r]; __Check((slice.Length).ToString(), "3");
