// vybe-test: csharp/csharp_patterns/multiple_classes_interacting
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var cart = new Cart();
cart.Add(new Item("Apple", 1.5));
cart.Add(new Item("Bread", 2.5));
cart.Add(new Item("Milk", 3.0));
__P((cart.Count()).ToString());
__P((cart.Total()).ToString());
__Check("3\n7");

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

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
