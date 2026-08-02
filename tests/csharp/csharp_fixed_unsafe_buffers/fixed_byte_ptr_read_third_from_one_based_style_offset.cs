// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_read_third_from_one_based_style_offset
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={2,4,6,8}; unsafe{fixed(byte* ptr=&arr[0]){__Check((ptr[2]).ToString(), "6");}}
