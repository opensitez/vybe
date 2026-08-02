// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_compare_elements_via_pointers
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={5,5,9}; unsafe{fixed(int* ptr=&arr[0]){__Check((ptr[0]==ptr[1]).ToString(), "True"); __Check((ptr[0]==ptr[2]).ToString(), "False");}}
