// vybe-test: csharp/csharp_deconstruction/foreach_deconstruction_reads_tuple_sequence
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

var pairs = new[] { ("a", 1), ("b", 2) };
foreach (var (letter, number) in pairs) {
    Console.WriteLine(letter + number);
}
