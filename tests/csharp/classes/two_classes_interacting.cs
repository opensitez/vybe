// vybe-test: csharp/classes/two_classes_interacting
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Engine {
            public int hp;
            public Engine(int h) { this.hp = h; }
        }
        class Car {
            public string model;
            public Engine engine;
            public Car(string m, int hp) {
                this.model = m;
                this.engine = new Engine(hp);
            }
        }
        var car = new Car("Tesla", 450);
        __Check((car.model).ToString(), "Tesla");
        __Check((car.engine.hp).ToString(), "450");
