// vybe-test: csharp/csharp_numerics_int128_uint128/uint128_max_value_boundary

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

UInt128 max = UInt128.MaxValue;
__P((max > 0).ToString());
__P(UInt128.MinValue.ToString());
__Check("True\n0");
