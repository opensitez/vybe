// vybe-test: csharp/csharp_generics_where/where_interface_constraint_calls_interface_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IName{string Name();}
class A:IName{public string Name()=>"A";}
string GetName<T>(T t) where T:IName=>t.Name();
__Check((GetName(new A())).ToString(), "A");
