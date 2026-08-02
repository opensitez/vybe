// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_nullable_struct_false_reads_zero
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={0,1}; unsafe{fixed(int* ptr=&arr[0]){__Check((*ptr==0).ToString(), "True");}}
