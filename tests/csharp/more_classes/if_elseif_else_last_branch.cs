// vybe-test: csharp/more_classes/if_elseif_else_last_branch
// origin: languages/csharp/tests/csharp/test_more_classes.rs

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

var x = 5;
        if (x > 20) { __P(("big").ToString()); }
        else if (x > 10) { __P(("medium").ToString()); }
        else { __P(("small").ToString()); }
__Check("small");
