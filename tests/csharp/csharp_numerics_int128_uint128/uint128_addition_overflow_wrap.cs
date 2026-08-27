// vybe-test: csharp/csharp_numerics_int128_uint128/uint128_addition_overflow_wrap

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

UInt128 a = (UInt128)100;
UInt128 b = (UInt128)200;
UInt128 sum = a + b;
__P(sum.ToString());
__Check("300");
