// vybe-test: csharp/interfaces_generics/generic_interface
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRepository<T> {
    void Add(T item);
    int Count();
}
class ListRepo<T> : IRepository<T> {
    private List<T> items = new List<T>();
    public void Add(T item) { items.Add(item); }
    public int Count() { return items.Count; }
}
var repo = new ListRepo<string>();
repo.Add("a");
repo.Add("b");
repo.Add("c");
__Check((repo.Count()).ToString(), "3");
