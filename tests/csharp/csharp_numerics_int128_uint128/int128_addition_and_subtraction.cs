// vybe-test: csharp/csharp_numerics_int128_uint128/int128_addition_and_subtraction

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

Int128 a = Int128.Parse("100000000000000000000");
Int128 b = Int128.Parse("200000000000000000000");
Int128 sum = a + b;
Int128 diff = b - a;
__P(sum.ToString());
__P(diff.ToString());
__Check("300000000000000000000\n100000000000000000000");
