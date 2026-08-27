// vybe-test: csharp/csharp_equality_contracts/string_equality_compares_character_sequence_not_reference_identity
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

using static __Harness;

string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
__P((a == b).ToString());
__P((object.ReferenceEquals(a, b)).ToString());
__Check("True\nFalse");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
