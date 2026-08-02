// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_assign_from_stack_value
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] arr={0}; byte temp=33; unsafe{fixed(byte* ptr=&arr[0]){*ptr=temp;}} __Check((arr[0]).ToString(), "33");
