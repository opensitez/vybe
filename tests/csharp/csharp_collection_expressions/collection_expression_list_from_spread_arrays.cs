// vybe-test: csharp/csharp_collection_expressions/collection_expression_list_from_spread_arrays
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

int[] a = [1, 2];
int[] b = [3];
System.Collections.Generic.List<int> list = [..a, ..b];
__P((list.Count).ToString()); __P((list[2]).ToString());
__Check("3\n3");
