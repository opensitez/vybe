// vybe-test: csharp/csharp_loops/goto_jumps_forward_to_labeled_statement
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

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

int x = 0;
goto done;
x = 99;
done:
__P((x).ToString());
__Check("0");
