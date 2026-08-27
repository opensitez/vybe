// vybe-test: csharp/csharp_numerics_int128_uint128/int128_min_max_constants

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

__P((Int128.MinValue < 0).ToString());
__P((Int128.MaxValue > 0).ToString());
__Check("True\nTrue");
