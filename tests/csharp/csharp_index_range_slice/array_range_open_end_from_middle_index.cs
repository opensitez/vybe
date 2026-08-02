// vybe-test: csharp/csharp_index_range_slice/array_range_open_end_from_middle_index
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={5,10,15,20,25}; var slice=data[2..]; __Check((slice.Length).ToString(), "3"); __Check((slice[0]).ToString(), "15");
