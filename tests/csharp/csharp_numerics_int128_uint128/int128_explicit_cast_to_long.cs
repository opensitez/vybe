// vybe-test: csharp/csharp_numerics_int128_uint128/int128_explicit_cast_to_long

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

Int128 v = (Int128)987654321L;
long l = (long)v;
__P(l.ToString());
__Check("987654321");
