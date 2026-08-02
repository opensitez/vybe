// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_xor_toggle_bit_in_place
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={0b1010}; unsafe{fixed(byte* ptr=&arr[0]){*ptr=(byte)(*ptr^0b1111);}} __Check((arr[0]).ToString(), "5");
