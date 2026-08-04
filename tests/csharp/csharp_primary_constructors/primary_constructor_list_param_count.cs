// vybe-test: csharp/csharp_primary_constructors/primary_constructor_list_param_count
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Bag(System.Collections.Generic.List<int> items) { public int Count => items.Count; }
__P((new Bag(new System.Collections.Generic.List<int> { 1, 2 }).Count).ToString());
__Check("2");
