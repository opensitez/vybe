// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_readonly_span_style_index_from_end
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2,3,4}; unsafe{fixed(int* ptr=&arr[0]){__Check((ptr[3]).ToString(), "4");}}
