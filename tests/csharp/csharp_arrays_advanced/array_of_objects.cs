// vybe-test: csharp/csharp_arrays_advanced/array_of_objects
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

class Item {
    public string Name;
    public Item(string n) { Name = n; }
}
var items = new[] { new Item("a"), new Item("b"), new Item("c") };
foreach (var item in items) {
    Console.WriteLine(item.Name);
}
