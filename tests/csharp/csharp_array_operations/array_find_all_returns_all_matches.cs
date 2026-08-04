// vybe-test: csharp/csharp_array_operations/array_find_all_returns_all_matches
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

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

int[] a = {1,2,3,4,5};
int[] evens = System.Array.FindAll(a, x => x%2==0);
__P((evens.Length).ToString());
__Check("2");
