// vybe-test: csharp/csharp_indexers/readonly_indexer_exposes_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

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

class Odds{public int this[int n]=>2*n+1;}
__P((new Odds()[4]).ToString());
__Check("9");
