// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_method_local_array
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Read(){int[] arr={42}; unsafe{fixed(int* ptr=&arr[0]){return *ptr;}} return 0;} __Check((Read()).ToString(), "42");
