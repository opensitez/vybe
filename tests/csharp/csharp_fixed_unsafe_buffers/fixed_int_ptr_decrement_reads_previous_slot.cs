// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_decrement_reads_previous_slot
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={3,4,5}; unsafe{fixed(int* ptr=&arr[2]){ptr--; __Check((*ptr).ToString(), "4");}}
