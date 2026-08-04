// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_with_outer_generic_constraint
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Repo<T>{public class Row{public T Data;} public T Read(Row r)=>r.Data;} __P((new Repo<int>().Read(new Repo<int>.Row{Data=77})).ToString());
__Check("77");
