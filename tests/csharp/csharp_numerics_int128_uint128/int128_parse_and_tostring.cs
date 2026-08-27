// vybe-test: csharp/csharp_numerics_int128_uint128/int128_parse_and_tostring

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

Int128 val = Int128.Parse("123456789012345678901234567890");
__P(val.ToString());
__Check("123456789012345678901234567890");
