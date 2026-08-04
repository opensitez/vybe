// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_empty_array_adds_nothing
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

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

int[] empty = [];
int[] all = [1, ..empty, 2];
__P((all.Length).ToString());
__Check("2");
