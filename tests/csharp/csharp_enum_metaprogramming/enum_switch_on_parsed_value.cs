// vybe-test: csharp/csharp_enum_metaprogramming/enum_switch_on_parsed_value
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

enum Mode{On,Off} var m=(Mode)System.Enum.Parse(typeof(Mode),"On"); string s=m==Mode.On?"yes":"no"; __P((s).ToString());
__Check("yes");
