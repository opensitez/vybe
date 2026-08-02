// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_from_last_index_reads_tail
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={100,200,300}; unsafe{fixed(int* ptr=&arr[2]){__Check((*ptr).ToString(), "300");}}
