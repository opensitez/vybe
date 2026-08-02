// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_swap_two_elements
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,9}; unsafe{fixed(int* ptr=&arr[0]){int t=ptr[0]; ptr[0]=ptr[1]; ptr[1]=t;}} __Check((arr[0]).ToString(), "9"); __Check((arr[1]).ToString(), "1");
