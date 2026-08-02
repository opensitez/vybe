// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_from_middle_index_reads_offset
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={10,20,30,40}; unsafe{fixed(byte* ptr=&arr[2]){__Check((*ptr).ToString(), "30");}}
