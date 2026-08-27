// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector2_distance_calculation

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

var p1 = new System.Numerics.Vector2(1.0f, 1.0f);
var p2 = new System.Numerics.Vector2(4.0f, 5.0f);
float dist = System.Numerics.Vector2.Distance(p1, p2);
__P(dist.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("5");
