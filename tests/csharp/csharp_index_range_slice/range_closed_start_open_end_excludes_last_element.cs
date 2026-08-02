// vybe-test: csharp/csharp_index_range_slice/range_closed_start_open_end_excludes_last_element
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={5,6,7,8,9}; var slice=data[1..^1]; __Check((slice.Length).ToString(), "3"); __Check((slice[0]).ToString(), "6"); __Check((slice[1]).ToString(), "7");
