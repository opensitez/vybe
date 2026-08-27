// vybe-test: csharp/csharp_numerics_half_precision_floats/half_tryparse_valid_and_invalid

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

bool ok1 = Half.TryParse("12.5", System.Globalization.CultureInfo.InvariantCulture, out Half res1);
bool ok2 = Half.TryParse("not_a_half", System.Globalization.CultureInfo.InvariantCulture, out Half res2);
__P(ok1.ToString());
__P(res1.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(ok2.ToString());
__Check("True\n12.5\nFalse");
