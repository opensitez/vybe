// vybe-test: csharp/csharp_enum_metaprogramming/enum_underlying_type_ushort
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

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

enum UShortEnum:ushort{Max=65535} __P((System.Enum.GetUnderlyingType(typeof(UShortEnum)).Name).ToString());
__Check("UInt16");
