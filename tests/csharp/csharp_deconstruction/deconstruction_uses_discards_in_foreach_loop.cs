// vybe-test: csharp/csharp_deconstruction/deconstruction_uses_discards_in_foreach_loop
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

var items = new[] { ("x", 1), ("y", 2), ("z", 3) };
foreach (var (_, number) in items) {
    Console.WriteLine(number * 10);
}
