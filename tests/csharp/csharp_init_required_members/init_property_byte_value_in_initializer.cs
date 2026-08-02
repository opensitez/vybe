// vybe-test: csharp/csharp_init_required_members/init_property_byte_value_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class ByteHolder { public byte Code { get; init; } }
var b = new ByteHolder { Code = 255 };
__Check((b.Code).ToString(), "255");
