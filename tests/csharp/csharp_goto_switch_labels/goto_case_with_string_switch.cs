// vybe-test: csharp/csharp_goto_switch_labels/goto_case_with_string_switch
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

string key = "b";
string r = "";
switch (key) {
    case "a": r += "A"; goto case "b";
    case "b": r += "B"; break;
    case "c": r += "C"; break;
}
__P((r).ToString());
__Check("B");
