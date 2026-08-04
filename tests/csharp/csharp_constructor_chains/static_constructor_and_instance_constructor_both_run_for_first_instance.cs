// vybe-test: csharp/csharp_constructor_chains/static_constructor_and_instance_constructor_both_run_for_first_instance
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

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

class Box { static Box() { __P(("static").ToString()); } public Box() { __P(("instance").ToString()); } } new Box();
__Check("static\ninstance");
