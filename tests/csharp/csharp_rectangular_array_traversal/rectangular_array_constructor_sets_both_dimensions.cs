// vybe-test: csharp/csharp_rectangular_array_traversal/rectangular_array_constructor_sets_both_dimensions
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

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

var grid = new int[2, 3];
__P((grid.GetLength(0)).ToString());
__P((grid.GetLength(1)).ToString());
__Check("2\n3");
