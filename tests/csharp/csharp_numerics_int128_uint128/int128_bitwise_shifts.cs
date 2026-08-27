// vybe-test: csharp/csharp_numerics_int128_uint128/int128_bitwise_shifts

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

Int128 a = (Int128)1;
Int128 shiftedLeft = a << 100;
Int128 shiftedRight = shiftedLeft >> 100;
__P((shiftedLeft > a).ToString());
__P((shiftedRight == a).ToString());
__Check("True\nTrue");
