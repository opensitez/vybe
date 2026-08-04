// vybe-test: csharp/csharp_new_features/target_typed_new_in_constructor_argument
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

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

class Box { public System.Collections.Generic.List<int> Items; public Box(System.Collections.Generic.List<int> i){Items=i;} }
var b = new Box(new());
b.Items.Add(9);
__P((b.Items.Count).ToString());
__Check("1");
