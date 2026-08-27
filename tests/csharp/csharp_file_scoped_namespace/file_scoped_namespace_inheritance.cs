// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

var d = new Dog();
__P((d.Name).ToString());
__P((d.Breed).ToString());
__Check("base\nlab");

class Animal { public string Name = "base"; }

class Dog : Animal { public string Breed = "lab"; }

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
