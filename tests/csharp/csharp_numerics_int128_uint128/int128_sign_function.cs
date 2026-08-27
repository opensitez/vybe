// vybe-test: csharp/csharp_numerics_int128_uint128/int128_sign_function

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

__P(Int128.Sign((Int128)(-42)).ToString());
__P(Int128.Sign((Int128)0).ToString());
__P(Int128.Sign((Int128)42).ToString());
__Check("-1\n0\n1");
