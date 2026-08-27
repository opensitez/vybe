// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_in_generic_list

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

var list = new System.Collections.Generic.List<System.Numerics.BigInteger>();
list.Add(System.Numerics.BigInteger.Parse("999999999999999999999"));
__P(list.Count.ToString());
__P(list[0].ToString());
__Check("1\n999999999999999999999");
