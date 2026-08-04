// vybe-test: csharp/csharp_object_initializers/list_of_objects_with_object_initializers
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

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

class Item{public int Id;}
var items=new System.Collections.Generic.List<Item>{new Item{Id=1},new Item{Id=2}};
__P((items[1].Id).ToString());
__Check("2");
