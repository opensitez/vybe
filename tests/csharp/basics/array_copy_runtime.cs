// vybe-test: csharp/basics/array_copy_runtime
// origin: languages/csharp/tests/csharp/test_basics.rs

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

int[] src = new int[] { 10, 20, 30, 40 };
int[] dst = new int[] { 0, 0, 0, 0 };
Array.Copy(src, dst, 3);
__P((dst[0] + dst[1] + dst[2] + dst[3]).ToString());
__Check("60");
