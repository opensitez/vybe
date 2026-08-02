// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_subtract_returns_to_first_element
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={11,22,33}; unsafe{fixed(int* ptr=&arr[1]){__Check((*(ptr-1)).ToString(), "11");}}
