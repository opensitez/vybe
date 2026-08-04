// vybe-test: csharp/csharp_reflection/property_info_set_value_writes_property_dynamically
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
var item = new Item();
var prop = typeof(Item).GetProperty("Id");
prop.SetValue(item, 99);
__P((item.Id).ToString());
__Check("99");
