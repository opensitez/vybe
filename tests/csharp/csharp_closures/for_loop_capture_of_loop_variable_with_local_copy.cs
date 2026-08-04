// vybe-test: csharp/csharp_closures/for_loop_capture_of_loop_variable_with_local_copy
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

var actions = new System.Collections.Generic.List<System.Func<int>>();
for(int i=0; i<3; i++) {
    var copy = i;
    actions.Add(() => copy);
}
foreach(var a in actions) __P((a()).ToString());
__Check("0\n1\n2");
