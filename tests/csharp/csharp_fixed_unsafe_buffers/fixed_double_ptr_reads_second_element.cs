// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_double_ptr_reads_second_element
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double[] arr={1.1,2.2}; unsafe{fixed(double* ptr=&arr[0]){__Check((*(ptr+1)).ToString(), "2.2");}}
