// vybe-test: csharp/csharp_numerics_half_precision_floats/half_get_type_code

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

TypeCode tc = Type.GetTypeCode(typeof(Half));
__P(tc.ToString());
__Check("Object");
