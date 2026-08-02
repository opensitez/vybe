// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_ushort_ptr_reads_unsigned_short
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

ushort[] arr={65000,1}; unsafe{fixed(ushort* ptr=&arr[0]){__Check((*ptr).ToString(), "65000");}}
