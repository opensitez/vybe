// vybe-test: csharp/oop_advanced/abstract_multiple_implementations
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((v.Type()).ToString(), "Car");
v = new Truck();
__Check((v.Type()).ToString(), "Truck");
