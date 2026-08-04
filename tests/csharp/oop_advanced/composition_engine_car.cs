// vybe-test: csharp/oop_advanced/composition_engine_car
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
var car = new Car("Sedan", 200);
__P((car.Info()).ToString());
__Check("Sedan 200hp");
