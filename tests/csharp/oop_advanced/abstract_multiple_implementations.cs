// vybe-test: csharp/oop_advanced/abstract_multiple_implementations
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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

abstract class Vehicle {
    public abstract string Type();
}
class Car : Vehicle {
    public override string Type() { return "Car"; }
}
class Truck : Vehicle {
    public override string Type() { return "Truck"; }
}
Vehicle v = new Car();
__P((v.Type()).ToString());
v = new Truck();
__P((v.Type()).ToString());
__Check("Car\nTruck");
