// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

// indexer_get_set
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
