// vybe-test: csharp/csharp_oop_inheritance/derived_class_can_have_additional_members
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

var bike = new Bike();
__P((bike.Wheels).ToString());
__P((bike.HasKickstand).ToString());
__Check("4\nTrue");

class Vehicle { public int Wheels = 4; }

class Bike : Vehicle { public bool HasKickstand = true; }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
