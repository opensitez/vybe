// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_write_updates_array_element
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){ptr[1]=99;}} __Check((arr[1]).ToString(), "99");
