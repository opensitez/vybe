// vybe-test: csharp/csharp_readonly_members/const_local_not_changeable_but_usable_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

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

const int MAX=100;
__P((MAX*2).ToString());
__Check("200");
