// vybe-test: csharp/csharp_patterns/multiple_classes_interacting
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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
Console.WriteLine(cart.Count());
Console.WriteLine(cart.Total());
