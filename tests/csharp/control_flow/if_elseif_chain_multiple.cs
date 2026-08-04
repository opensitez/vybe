// vybe-test: csharp/control_flow/if_elseif_chain_multiple
// origin: languages/csharp/tests/csharp/test_control_flow.rs

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

var x = 2;
        if (x == 1) { __P(("one").ToString()); }
        else if (x == 2) { __P(("two").ToString()); }
        else if (x == 3) { __P(("three").ToString()); }
        else { __P(("other").ToString()); }
__Check("two");
