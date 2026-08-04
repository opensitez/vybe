// vybe-test: csharp/csharp_collections_initialise/collection_expression_spread_merges_two_spans
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

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

int[] a=[1,2,3];
int[] b=[4,5,6];
int[] c=[..a,..b];
__P((c.Length).ToString()); __P((c[3]).ToString());
__Check("6\n4");
