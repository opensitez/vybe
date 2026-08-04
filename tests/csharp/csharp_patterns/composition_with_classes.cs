// vybe-test: csharp/csharp_patterns/composition_with_classes
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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
var car = new Car("Toyota", 200);
__P((car.Describe()).ToString());
__Check("Toyota 200hp");
