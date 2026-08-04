// vybe-test: csharp/csharp_array_2d/two_d_array_initializer_syntax
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

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

int[,] m={{1,2,3},{4,5,6}};
__P((m[0,2]).ToString()); __P((m[1,0]).ToString());
__Check("3\n4");
