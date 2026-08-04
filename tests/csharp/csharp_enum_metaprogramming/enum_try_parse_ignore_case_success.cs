// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_ignore_case_success
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

enum Mode{Alpha,Beta} var ok=System.Enum.TryParse<Mode>("beta",true,out var m); __P((ok).ToString()); __P((m).ToString());
__Check("True\nBeta");
