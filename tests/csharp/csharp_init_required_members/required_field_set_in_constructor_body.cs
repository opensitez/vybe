// vybe-test: csharp/csharp_init_required_members/required_field_set_in_constructor_body
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Packet { public required int Size; public Packet(int size) { Size = size; } }
var p = new Packet(256);
__Check((p.Size).ToString(), "256");
