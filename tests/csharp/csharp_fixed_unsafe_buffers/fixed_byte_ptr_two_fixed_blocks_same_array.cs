// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_two_fixed_blocks_same_array
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={1,2,3}; unsafe{fixed(byte* a=&arr[0]){fixed(byte* b=&arr[1]){__Check((*a+b[0]).ToString(), "3");}}}
