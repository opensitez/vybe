// vybe-test: csharp/csharp_pattern_list/is_list_byte_array_length_two_discard
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] buf=new byte[]{10,20}; __Check((buf is [_,_]).ToString(), "True");
