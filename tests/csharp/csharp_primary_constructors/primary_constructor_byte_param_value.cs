// vybe-test: csharp/csharp_primary_constructors/primary_constructor_byte_param_value
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class ByteBox(byte b) { public byte Value => b; }
__Check((new ByteBox(200).Value).ToString(), "200");
