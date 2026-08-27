// vybe-test: csharp/csharp_numerics_int128_uint128/int128_is_negative_and_is_even

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

Int128 neg = (Int128)(-10);
Int128 evenVal = (Int128)20;
__P(Int128.IsNegative(neg).ToString());
__P(Int128.IsEvenInteger(evenVal).ToString());
__Check("True\nTrue");
