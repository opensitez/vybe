// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_dereference_write_mutates_source
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={4,5,6}; unsafe{fixed(int* ptr=&arr[0]){*(ptr+1)=88;}} __Check((arr[1]).ToString(), "88");
