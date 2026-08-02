// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_from_local_array_reads_backing_store
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unsafe{byte[] arr={5,6}; fixed(byte* ptr=&arr[0]){__Check((ptr[1]).ToString(), "6");}}
