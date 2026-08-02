// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_address_of_element_equals_base_plus_offset
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={10,20,30}; unsafe{fixed(int* basePtr=&arr[0]){fixed(int* off=&arr[2]){__Check((off-basePtr).ToString(), "2");}}}
