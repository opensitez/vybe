// vybe-test: csharp/csharp_index_range_slice/range_start_from_end_end_open
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4}; System.Range r=^2..; var slice=data[r]; __Check((slice[0]).ToString(), "3"); __Check((slice[1]).ToString(), "4");
