// vybe-test: csharp/strings_advanced/convert_toint32_tostring
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

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

int x = Convert.ToInt32("123");
string s = Convert.ToString(456);
__P((x).ToString());
__P((s).ToString());
__Check("123\n456");
