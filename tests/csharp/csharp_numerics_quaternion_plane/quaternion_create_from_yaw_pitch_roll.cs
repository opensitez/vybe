// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_create_from_yaw_pitch_roll

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

var q = System.Numerics.Quaternion.CreateFromYawPitchRoll(0.0f, 0.0f, 0.0f);
__P(q.IsIdentity.ToString());
__Check("True");
