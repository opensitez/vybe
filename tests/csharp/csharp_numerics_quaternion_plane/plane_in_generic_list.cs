// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_in_generic_list

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

var list = new System.Collections.Generic.List<System.Numerics.Plane>();
list.Add(new System.Numerics.Plane(System.Numerics.Vector3.UnitZ, 0));
__P(list.Count.ToString());
__Check("1");
