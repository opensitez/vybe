// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_from_field_backed_array
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder{public byte[] Data={9,8};} var h=new Holder(); unsafe{fixed(byte* ptr=&h.Data[0]){__Check((*ptr).ToString(), "9");}}
