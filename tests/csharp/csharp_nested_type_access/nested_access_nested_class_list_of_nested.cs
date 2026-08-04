// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_list_of_nested
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

class Bag{public class Item{public int Id;} public System.Collections.Generic.List<Item> All(){var list=new System.Collections.Generic.List<Item>(); list.Add(new Item{Id=1}); return list;}} __P((new Bag().All()[0].Id).ToString());
__Check("1");
