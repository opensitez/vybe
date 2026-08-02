// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_xor_toggles_bits_present_in_one_operand
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags]
enum Perm { A = 1, B = 2 }
var value = (Perm.A | Perm.B) ^ Perm.A;
__Check(((int)value).ToString(), "2");
