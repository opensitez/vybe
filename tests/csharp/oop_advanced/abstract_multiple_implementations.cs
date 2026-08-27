// vybe-test: csharp/oop_advanced/abstract_multiple_implementations
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

Vehicle v = new Car();
__P((v.Type()).ToString());
v = new Truck();
__P((v.Type()).ToString());
__Check("Car\nTruck");

abstract class Vehicle {
    public abstract string Type();
}

class Car : Vehicle {
    public override string Type() { return "Car"; }
}

class Truck : Vehicle {
    public override string Type() { return "Truck"; }
}

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
