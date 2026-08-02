// vybe-test: csharp/csharp_enum_operations/enum_with_explicit_underlying_byte_type
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Small:byte{A=1,B=200}
__Check(((byte)Small.B).ToString(), "200");
