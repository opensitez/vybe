// vybe-test: csharp/csharp_closures/nested_closure_captures_from_outer_scope
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

System.Func<int,System.Func<int>> makeAdder = x => () => x + 1;
var add1 = makeAdder(5);
__P((add1()).ToString());
__Check("6");
