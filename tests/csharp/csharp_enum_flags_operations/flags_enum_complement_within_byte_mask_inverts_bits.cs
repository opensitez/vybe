// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_complement_within_byte_mask_inverts_bits
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags]
enum Perm : byte { A = 1, B = 2 }
var value = Perm.A | Perm.B;
var cleared = value & ~Perm.A;
__Check(((int)cleared).ToString(), "2");
