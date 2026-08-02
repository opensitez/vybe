// vybe-test: csharp/oop_advanced/composition_engine_car
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((car.Info()).ToString(), "Sedan 200hp");
