// vybe-test: csharp/csharp_numerics_int128_uint128/int128_abs_function

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

Int128 neg = (Int128)(-5000);
Int128 abs = Int128.Abs(neg);
__P(abs.ToString());
__Check("5000");
