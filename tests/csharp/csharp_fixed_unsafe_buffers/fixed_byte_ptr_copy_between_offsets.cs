// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_copy_between_offsets
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){ptr[2]=ptr[0];}} __Check((arr[2]).ToString(), "1");
