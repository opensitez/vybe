// vybe-test: csharp/csharp_interface_contracts/ienumerable_foreach_drives_custom_iterator
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

class Counter : System.Collections.Generic.IEnumerable<int> {
    public System.Collections.Generic.IEnumerator<int> GetEnumerator() {
        yield return 1; yield return 2; yield return 3;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
int sum=0;
foreach(var n in new Counter()) sum+=n;
Console.WriteLine(sum);
