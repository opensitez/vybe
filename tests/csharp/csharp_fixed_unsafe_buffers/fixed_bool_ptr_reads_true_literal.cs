// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_bool_ptr_reads_true_literal
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool[] arr={true,false}; unsafe{fixed(bool* ptr=&arr[0]){__Check((*ptr).ToString(), "True");}}
