// vybe-test: csharp/csharp_oop_inheritance/derived_class_can_have_additional_members
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Vehicle { public int Wheels = 4; }
class Bike : Vehicle { public bool HasKickstand = true; }
var bike = new Bike();
__P((bike.Wheels).ToString());
__P((bike.HasKickstand).ToString());
__Check("4\nTrue");
