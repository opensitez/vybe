// vybe-test: csharp/csharp_null_handling/null_conditional_indexer_returns_null_on_null_collection
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

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

int[] arr = null;
__P((arr?[0] ?? -1).ToString());
__Check("-1");
