// vybe-test: csharp/oop_advanced/composition_engine_car
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var car = new Car("Sedan", 200);
__P((car.Info()).ToString());
__Check("Sedan 200hp");

class Engine {
    public int Horsepower { get; set; }
    public Engine(int hp) { Horsepower = hp; }
}

class Car {
    public string Name { get; set; }
    public Engine Engine { get; set; }
    public Car(string name, int hp) {
        Name = name;
        Engine = new Engine(hp);
    }
    public string Info() { return Name + " " + Engine.Horsepower + "hp"; }
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
