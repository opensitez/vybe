// vybe-test: csharp/csharp_deconstruction_patterns/deconstruction_assignment_to_existing_variables
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

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

int x=0, y=0;
(x, y) = (5, 10);
__P((x).ToString()); __P((y).ToString());
__Check("5\n10");
