// vybe-test: csharp/csharp_numerics_int128_uint128/int128_tryparse_valid_and_invalid

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

bool ok1 = Int128.TryParse("987654321", out Int128 v1);
bool ok2 = Int128.TryParse("invalid_number", out Int128 v2);
__P(ok1.ToString());
__P(v1.ToString());
__P(ok2.ToString());
__Check("True\n987654321\nFalse");
