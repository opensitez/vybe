// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_nested_unsafe_reads_value
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={9}; unsafe{fixed(int* outer=&arr[0]){unsafe{__Check((*outer).ToString(), "9");}}}
