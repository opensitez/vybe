// vybe-test: csharp/csharp_numerics_int128_uint128/int128_icomparable_ordering

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

Int128 v1 = (Int128)100;
Int128 v2 = (Int128)200;
__P((v1.CompareTo(v2) < 0).ToString());
__Check("True");
