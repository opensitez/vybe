// vybe-test: csharp/csharp_collection_expressions/collection_expression_nested_spread_flatten
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

int[] a = [1]; int[] b = [2]; int[] c = [3];
int[] all = [..a, ..b, ..c];
__P((string.Join(",", all)).ToString());
__Check("1,2,3");
