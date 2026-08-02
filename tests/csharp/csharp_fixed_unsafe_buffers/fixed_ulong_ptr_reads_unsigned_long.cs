// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_ulong_ptr_reads_unsigned_long
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

ulong[] arr={18446744073709551615UL,0UL}; unsafe{fixed(ulong* ptr=&arr[1]){__Check((*ptr).ToString(), "0");}}
