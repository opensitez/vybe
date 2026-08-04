// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_indexer_reads_values_through_interface
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

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

interface IReadIndex { string this[int index] { get; } }
class Words : IReadIndex {
    string[] values = new[] { "alpha", "beta" };
    string IReadIndex.this[int index] { get { return values[index]; } }
}
IReadIndex words = new Words();
__P((words[0]).ToString());
__P((words[1]).ToString());
__Check("alpha\nbeta");
