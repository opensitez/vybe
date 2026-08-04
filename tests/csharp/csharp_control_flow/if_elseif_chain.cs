// vybe-test: csharp/csharp_control_flow/if_elseif_chain
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

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

int score = 75;
if (score >= 90) __P(("A").ToString());
else if (score >= 80) __P(("B").ToString());
else if (score >= 70) __P(("C").ToString());
else __P(("F").ToString());
__Check("C");
