// vybe-test: csharp/csharp_constructor_chains/constructor_can_initialize_collection_field
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

using System.Collections.Generic; class Box { List<int> values; public Box() { values = new List<int> { 1, 2, 3 }; } public int Count() { return values.Count; } } __P((new Box().Count()).ToString());
__Check("3");
