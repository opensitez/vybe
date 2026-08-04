// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_at_start
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

int[] tail = [3, 4];
int[] all = [..tail, 1, 2];
__P((all[0]).ToString()); __P((all[2]).ToString());
__Check("3\n1");
