// vybe-test: csharp/csharp_init_required_members/required_field_set_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Packet { public required int Size; }
var p = new Packet { Size = 512 };
__Check((p.Size).ToString(), "512");
