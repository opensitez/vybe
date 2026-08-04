// vybe-test: csharp/csharp_equality_contracts/string_equality_compares_character_sequence_not_reference_identity
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
__P((a == b).ToString());
__P((object.ReferenceEquals(a, b)).ToString());
__Check("True\nFalse");
