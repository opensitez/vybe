// vybe-test: csharp/csharp_numerics_int128_uint128/int128_multiplication_and_division

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

Int128 a = Int128.Parse("1000000000000");
Int128 b = Int128.Parse("2000000000000");
Int128 prod = a * b;
Int128 div = prod / a;
__P(prod.ToString());
__P(div.ToString());
__Check("2000000000000000000000000\n2000000000000");
