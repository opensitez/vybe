// vybe-test: csharp/csharp_generics_where/multiple_where_constraints_combined
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

using static __Harness;

T Create<T>() where T:IGreet,new()=>new T();
__P((Create<Person>().Hi()).ToString());
__Check("hello");

interface IGreet{string Hi();}

class Person:IGreet{public string Hi()=>"hello"; public Person(){}}

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
