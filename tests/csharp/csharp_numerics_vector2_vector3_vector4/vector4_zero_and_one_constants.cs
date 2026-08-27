// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector4_zero_and_one_constants

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

__P((System.Numerics.Vector4.Zero.X == 0.0f).ToString());
__P((System.Numerics.Vector4.One.X == 1.0f).ToString());
__Check("True\nTrue");
