// vybe-test: csharp/interfaces_generics/generic_interface
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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
__P((repo.Count()).ToString());
__Check("3");
