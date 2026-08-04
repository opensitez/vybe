// vybe-test: csharp/csharp_array_2d/two_d_array_set_and_get_by_index
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

int[,] m=new int[3,3];
m[1,2]=99;
__P((m[1,2]).ToString());
__Check("99");
