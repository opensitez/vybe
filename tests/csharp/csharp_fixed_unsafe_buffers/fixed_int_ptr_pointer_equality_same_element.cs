// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_pointer_equality_same_element
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2}; unsafe{fixed(int* a=&arr[0]){fixed(int* b=&arr[0]){__Check((a==b).ToString(), "True");}}}
