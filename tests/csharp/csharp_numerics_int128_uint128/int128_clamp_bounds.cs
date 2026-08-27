// vybe-test: csharp/csharp_numerics_int128_uint128/int128_clamp_bounds

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

Int128 val = (Int128)150;
Int128 min = (Int128)0;
Int128 max = (Int128)100;
Int128 clamped = Int128.Clamp(val, min, max);
__P(clamped.ToString());
__Check("100");
