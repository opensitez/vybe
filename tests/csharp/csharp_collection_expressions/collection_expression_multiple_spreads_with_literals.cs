// vybe-test: csharp/csharp_collection_expressions/collection_expression_multiple_spreads_with_literals
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

int[] a = [1, 2]; int[] b = [3];
int[] c = [0, ..a, ..b, 4];
__P((c[0]).ToString()); __P((c[4]).ToString());
__Check("0\n4");
