// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_char_ptr_write_updates_string_builder_backing
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] arr={'x','y'}; unsafe{fixed(char* ptr=&arr[0]){ptr[1]='z';}} __Check((arr[1]).ToString(), "122");
