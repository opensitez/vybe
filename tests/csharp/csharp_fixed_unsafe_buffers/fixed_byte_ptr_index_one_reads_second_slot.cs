// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_index_one_reads_second_slot
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){__Check((ptr[1]).ToString(), "2");}}
