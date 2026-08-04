// vybe-test: csharp/csharp_struct_copy_independence/struct_field_mutation_on_copy_leaves_original_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_struct_copy_independence.rs

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

struct Point { public int X; }
var left = new Point { X = 1 };
var right = left;
right.X = 9;
__P((left.X).ToString());
__P((right.X).ToString());
__Check("1\n9");
