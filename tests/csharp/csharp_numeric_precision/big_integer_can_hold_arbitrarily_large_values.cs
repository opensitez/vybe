// vybe-test: csharp/csharp_numeric_precision/big_integer_can_hold_arbitrarily_large_values
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

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

var n=System.Numerics.BigInteger.Pow(10,30);
__P((n.ToString().StartsWith("1")).ToString());
__Check("True");
