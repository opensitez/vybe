// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_distance_between_elements_is_one
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2,3}; unsafe{fixed(int* a=&arr[0]){fixed(int* b=&arr[1]){__Check((b-a).ToString(), "1");}}}
