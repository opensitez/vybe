// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_and_masks_to_intersection_of_bits
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags]
enum Perm { A = 1, B = 2, C = 4 }
var combined = Perm.A | Perm.B | Perm.C;
var masked = combined & Perm.B;
__Check(((int)masked).ToString(), "2");
