// vybe-test: csharp/more_classes/auto_property_set_after_construction
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item {
            public string Name { get; set; }
            public Item() {}
        }
        var item = new Item();
        item.Name = "Widget";
        __Check((item.Name).ToString(), "Widget");
