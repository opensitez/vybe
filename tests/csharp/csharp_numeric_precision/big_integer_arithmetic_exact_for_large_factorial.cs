// vybe-test: csharp/csharp_numeric_precision/big_integer_arithmetic_exact_for_large_factorial
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

System.Numerics.BigInteger f=1;
for(int i=1;i<=20;i++) f*=i;
__P((f.ToString()).ToString());
__Check("2432902008176640000");
