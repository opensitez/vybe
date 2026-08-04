// vybe-test: csharp/csharp_collections/array_creation
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

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

int[] arr = {5, 10, 15, 20, 25};
__P((arr.Length).ToString());
__P((arr[2]).ToString());
__Check("5\n15");
