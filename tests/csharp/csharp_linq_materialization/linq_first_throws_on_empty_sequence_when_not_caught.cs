// vybe-test: csharp/csharp_linq_materialization/linq_first_throws_on_empty_sequence_when_not_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

using System.Linq;
try {
    Console.WriteLine(new int[0].First());
} catch (System.InvalidOperationException) {
    Console.WriteLine("empty");
}
