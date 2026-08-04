// vybe-test: csharp/basics/multi_var_three_variables
// origin: languages/csharp/tests/csharp/test_basics.rs

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

int x = 10, y = 20, z = 30;
        __P((x + y + z).ToString());
__Check("60");
