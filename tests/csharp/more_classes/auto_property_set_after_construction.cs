// vybe-test: csharp/more_classes/auto_property_set_after_construction
// origin: languages/csharp/tests/csharp/test_more_classes.rs

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

class Item {
            public string Name { get; set; }
            public Item() {}
        }
        var item = new Item();
        item.Name = "Widget";
        __P((item.Name).ToString());
__Check("Widget");
