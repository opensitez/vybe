// vybe-test: csharp/csharp_patterns/composition_with_classes
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((car.Describe()).ToString(), "Toyota 200hp");
