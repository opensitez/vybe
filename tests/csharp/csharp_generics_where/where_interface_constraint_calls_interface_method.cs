// vybe-test: csharp/csharp_generics_where/where_interface_constraint_calls_interface_method
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

interface IName{string Name();}
class A:IName{public string Name()=>"A";}
string GetName<T>(T t) where T:IName=>t.Name();
__P((GetName(new A())).ToString());
__Check("A");
