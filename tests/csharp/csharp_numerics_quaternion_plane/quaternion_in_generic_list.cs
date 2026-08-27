// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_in_generic_list

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

var list = new System.Collections.Generic.List<System.Numerics.Quaternion>();
list.Add(System.Numerics.Quaternion.Identity);
__P(list.Count.ToString());
__Check("1");
