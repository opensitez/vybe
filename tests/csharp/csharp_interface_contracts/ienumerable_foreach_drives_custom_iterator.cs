// vybe-test: csharp/csharp_interface_contracts/ienumerable_foreach_drives_custom_iterator
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

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

class Counter : System.Collections.Generic.IEnumerable<int> {
    public System.Collections.Generic.IEnumerator<int> GetEnumerator() {
        yield return 1; yield return 2; yield return 3;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
int sum=0;
foreach(var n in new Counter()) sum+=n;
__P((sum).ToString());
__Check("6");
