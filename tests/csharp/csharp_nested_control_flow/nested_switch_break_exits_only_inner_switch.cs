// vybe-test: csharp/csharp_nested_control_flow/nested_switch_break_exits_only_inner_switch
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

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

string report = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            switch (i) {
                case 0:
                    report += "inner;";
                    break;
            }
            report += "after-inner;";
            break;
        case 1:
            report += "tail;";
            break;
    }
}
__P((report).ToString());
__Check("inner;after-inner;tail;");
