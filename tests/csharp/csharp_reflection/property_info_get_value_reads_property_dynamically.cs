// vybe-test: csharp/csharp_reflection/property_info_get_value_reads_property_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

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

class Item { public int Id {get;set;} }
var item = new Item { Id=7 };
var prop = typeof(Item).GetProperty("Id");
__P((prop.GetValue(item)).ToString());
__Check("7");
