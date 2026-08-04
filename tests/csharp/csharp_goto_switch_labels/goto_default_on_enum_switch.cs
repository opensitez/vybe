// vybe-test: csharp/csharp_goto_switch_labels/goto_default_on_enum_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

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

enum Color { Red, Green }
Color c = (Color)9;
string name = "";
switch (c) {
    case Color.Red: name = "R"; break;
    case Color.Green: name = "G"; break;
    default: name = "?"; break;
}
__P((name).ToString());
__Check("?");
