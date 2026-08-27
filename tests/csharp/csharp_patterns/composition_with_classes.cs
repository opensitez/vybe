// vybe-test: csharp/csharp_patterns/composition_with_classes
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var car = new Car("Toyota", 200);
__P((car.Describe()).ToString());
__Check("Toyota 200hp");

class Engine {
    public int Horsepower;
    public Engine(int hp) { Horsepower = hp; }
}

class Car {
    public string Make;
    public Engine Engine;
    public Car(string make, int hp) {
        Make = make;
        Engine = new Engine(hp);
    }
    public string Describe() {
        return Make + " " + Engine.Horsepower + "hp";
    }
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
