// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_scale_offset_by_element_size
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={10,20,30,40}; unsafe{fixed(int* ptr=&arr[0]){__Check((*(ptr+3)).ToString(), "40");}}
