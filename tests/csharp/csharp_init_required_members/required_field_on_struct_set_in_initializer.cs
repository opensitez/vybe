// vybe-test: csharp/csharp_init_required_members/required_field_on_struct_set_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Block { public required int Length; }
var b = new Block { Length = 64 };
__Check((b.Length).ToString(), "64");
