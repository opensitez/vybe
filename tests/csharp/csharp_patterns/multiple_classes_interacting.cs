// vybe-test: csharp/csharp_patterns/multiple_classes_interacting
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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
    public string Name;
    public double Price;
    public Item(string n, double p) { Name = n; Price = p; }
}
class Cart {
    private List<Item> items = new List<Item>();
    public void Add(Item item) { items.Add(item); }
    public int Count() { return items.Count; }
    public double Total() {
        double sum = 0;
        foreach (var item in items) sum += item.Price;
        return sum;
    }
}
var cart = new Cart();
cart.Add(new Item("Apple", 1.5));
cart.Add(new Item("Bread", 2.5));
cart.Add(new Item("Milk", 3.0));
__P((cart.Count()).ToString());
__P((cart.Total()).ToString());
__Check("3\n7");
