// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_uint_ptr_reads_unsigned_int
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

uint[] arr={3000000000u,1u}; unsafe{fixed(uint* ptr=&arr[0]){__Check((*ptr>0).ToString(), "True");}}
