// vybe-test: csharp/csharp_generics_where/multiple_where_constraints_combined
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

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

interface IGreet{string Hi();}
class Person:IGreet{public string Hi()=>"hello"; public Person(){}}
T Create<T>() where T:IGreet,new()=>new T();
__P((Create<Person>().Hi()).ToString());
__Check("hello");
