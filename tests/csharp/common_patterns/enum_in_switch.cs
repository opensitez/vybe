// vybe-test: csharp/common_patterns/enum_in_switch
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

enum Season { Spring, Summer, Autumn, Winter }
Season s = Season.Summer;
switch (s) {
    case Season.Spring: __P(("spring").ToString()); break;
    case Season.Summer: __P(("summer").ToString()); break;
    case Season.Autumn: __P(("autumn").ToString()); break;
    case Season.Winter: __P(("winter").ToString()); break;
}
__Check("summer");
