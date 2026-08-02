// vybe-test: csharp/csharp_interface_contracts/icomparable_implementation_used_by_list_sort
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

class Priority : System.IComparable<Priority> {
    public int Level;
    public int CompareTo(Priority other) => Level.CompareTo(other.Level);
}
var list = new System.Collections.Generic.List<Priority> {
    new Priority{Level=3}, new Priority{Level=1}, new Priority{Level=2}
};
list.Sort();
foreach(var p in list) Console.WriteLine(p.Level);
