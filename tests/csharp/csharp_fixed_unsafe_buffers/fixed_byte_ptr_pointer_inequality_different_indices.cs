// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_pointer_inequality_different_indices
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={1,2}; unsafe{fixed(byte* a=&arr[0]){fixed(byte* b=&arr[1]){__Check((a==b).ToString(), "False");}}}
