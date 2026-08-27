// vybe-test: csharp/csharp_classes/class_basic
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var p = new Person("Alice", 30);
__P((p.Describe()).ToString());
__Check("Alice is 30");

class Person {
    public string Name;
    public int Age;
    public Person(string name, int age) {
        Name = name;
        Age = age;
    }
    public string Describe() {
        return Name + " is " + Age;
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
