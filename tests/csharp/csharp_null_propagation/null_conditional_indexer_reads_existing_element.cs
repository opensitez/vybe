// vybe-test: csharp/csharp_null_propagation/null_conditional_indexer_reads_existing_element
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

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

int[] values = { 3, 4, 5 }; __P((values?[1] ?? -1).ToString());
__Check("4");
