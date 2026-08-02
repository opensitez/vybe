// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_float_ptr_reads_fractional_value
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

float[] arr={1.5f,2.5f}; unsafe{fixed(float* ptr=&arr[0]){__Check((*ptr==1.5f).ToString(), "True");}}
