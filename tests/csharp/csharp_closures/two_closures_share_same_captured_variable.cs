// vybe-test: csharp/csharp_closures/two_closures_share_same_captured_variable
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

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

int shared = 0;
System.Action add = () => shared++;
System.Func<int> read = () => shared;
add(); add();
__P((read()).ToString());
__Check("2");
