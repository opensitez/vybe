// vybe-test: csharp/csharp_generics_where/multiple_where_constraints_combined
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGreet{string Hi();}
class Person:IGreet{public string Hi()=>"hello"; public Person(){}}
T Create<T>() where T:IGreet,new()=>new T();
__Check((Create<Person>().Hi()).ToString(), "hello");
