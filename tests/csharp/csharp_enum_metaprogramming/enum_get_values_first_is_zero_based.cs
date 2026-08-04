// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_first_is_zero_based
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

enum Rank{First,Second} foreach(var v in System.Enum.GetValues(typeof(Rank))) __P(((int)v).ToString());
__Check("0\n1");
