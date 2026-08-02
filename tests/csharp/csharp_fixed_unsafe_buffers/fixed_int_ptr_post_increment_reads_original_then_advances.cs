// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_post_increment_reads_original_then_advances
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={2,3}; unsafe{fixed(int* ptr=&arr[0]){__Check((*ptr++).ToString(), "2"); __Check((*ptr).ToString(), "3");}}
