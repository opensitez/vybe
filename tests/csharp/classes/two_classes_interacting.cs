// vybe-test: csharp/classes/two_classes_interacting
// origin: languages/csharp/tests/csharp/test_classes.rs

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
        __P((car.model).ToString());
        __P((car.engine.hp).ToString());
__Check("Tesla\n450");
