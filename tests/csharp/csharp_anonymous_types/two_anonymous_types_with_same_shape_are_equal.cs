// vybe-test: csharp/csharp_anonymous_types/two_anonymous_types_with_same_shape_are_equal
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

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

var a=new{X=1,Y=2}; var b=new{X=1,Y=2};
__P((a.Equals(b)).ToString());
__Check("True");
