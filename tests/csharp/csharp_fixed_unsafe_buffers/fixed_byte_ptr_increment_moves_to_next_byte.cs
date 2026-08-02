// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_increment_moves_to_next_byte
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={7,8,9}; unsafe{fixed(byte* ptr=&arr[0]){ptr++; __Check((*ptr).ToString(), "8");}}
