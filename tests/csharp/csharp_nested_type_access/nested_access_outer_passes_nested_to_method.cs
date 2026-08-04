// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_passes_nested_to_method
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

class Store{public class Item{public int Id;} int Inspect(Item i)=>i.Id; public int Check()=>Inspect(new Item{Id=44});} __P((new Store().Check()).ToString());
__Check("44");
