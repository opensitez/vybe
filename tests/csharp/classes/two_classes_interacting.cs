// vybe-test: csharp/classes/two_classes_interacting
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var car = new Car("Tesla", 450);
__P((car.model).ToString());
__P((car.engine.hp).ToString());
__Check("Tesla\n450");

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
