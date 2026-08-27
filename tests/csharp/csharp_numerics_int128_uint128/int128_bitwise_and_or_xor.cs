// vybe-test: csharp/csharp_numerics_int128_uint128/int128_bitwise_and_or_xor

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

Int128 a = (Int128)0xFF00;
Int128 b = (Int128)0x0FF0;
__P(((a & b) == (Int128)0x0F00).ToString());
__P(((a | b) == (Int128)0xFFF0).ToString());
__P(((a ^ b) == (Int128)0xF0F0).ToString());
__Check("True\nTrue\nTrue");
