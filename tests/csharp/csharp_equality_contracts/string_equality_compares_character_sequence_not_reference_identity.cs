// vybe-test: csharp/csharp_equality_contracts/string_equality_compares_character_sequence_not_reference_identity
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
__Check((a == b).ToString(), "True");
__Check((object.ReferenceEquals(a, b)).ToString(), "False");
