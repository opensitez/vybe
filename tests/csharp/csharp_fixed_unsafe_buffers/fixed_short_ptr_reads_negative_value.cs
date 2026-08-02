// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_short_ptr_reads_negative_value
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

short[] arr={-3,4}; unsafe{fixed(short* ptr=&arr[0]){__Check((*ptr).ToString(), "-3");}}
