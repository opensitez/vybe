// vybe-test: csharp/csharp_index_range_slice/range_variable_used_for_array_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={2,4,6,8}; System.Range r=new System.Range(1,3); var slice=data[r]; __Check((slice.Length).ToString(), "2"); __Check((slice[1]).ToString(), "6");
