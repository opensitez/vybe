// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_long_ptr_reads_large_literal
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long[] arr={10000000000L,2L}; unsafe{fixed(long* ptr=&arr[0]){__Check((*ptr>0).ToString(), "True");}}
