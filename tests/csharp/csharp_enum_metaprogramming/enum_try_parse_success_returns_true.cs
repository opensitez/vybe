// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_success_returns_true
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

enum Day{Mon,Tue,Wed} var ok=System.Enum.TryParse<Day>("Tue",out var d); __P((ok).ToString()); __P((d).ToString());
__Check("True\nTue");
