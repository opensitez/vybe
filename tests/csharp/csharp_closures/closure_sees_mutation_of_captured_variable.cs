// vybe-test: csharp/csharp_closures/closure_sees_mutation_of_captured_variable
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

int x = 0;
System.Func<int> read = () => x;
x = 99;
__P((read()).ToString());
__Check("99");
