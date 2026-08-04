// vybe-test: csharp/csharp_struct_features/struct_assignment_copies_all_fields_independently
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

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

struct Point { public int X, Y; }
var a = new Point { X=1, Y=2 };
var b = a;
b.X = 99;
__P((a.X).ToString());
__Check("1");
