// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_factorial_computation

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

System.Numerics.BigInteger fact = 1;
for (int i = 1; i <= 10; i++) {
    fact *= i;
}
__P(fact.ToString());
__Check("3628800");
