// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_sbyte_ptr_reads_signed_byte
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

sbyte[] arr={-1,1}; unsafe{fixed(sbyte* ptr=&arr[0]){__Check((*ptr).ToString(), "-1");}}
