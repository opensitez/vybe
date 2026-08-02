// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_max_byte_value_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={255,0}; unsafe{fixed(byte* ptr=&arr[0]){__Check((*ptr).ToString(), "255");}}
