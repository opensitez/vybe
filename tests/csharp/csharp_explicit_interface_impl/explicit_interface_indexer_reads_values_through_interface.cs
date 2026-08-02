// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_indexer_reads_values_through_interface
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IReadIndex { string this[int index] { get; } }
class Words : IReadIndex {
    string[] values = new[] { "alpha", "beta" };
    string IReadIndex.this[int index] { get { return values[index]; } }
}
IReadIndex words = new Words();
__Check((words[0]).ToString(), "alpha");
__Check((words[1]).ToString(), "beta");
