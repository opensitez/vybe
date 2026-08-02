// vybe-test: csharp/csharp_index_range_slice/range_on_byte_array_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] data={10,20,30,40}; var slice=data[1..3]; __Check((slice[0]).ToString(), "20"); __Check((slice[1]).ToString(), "30");
