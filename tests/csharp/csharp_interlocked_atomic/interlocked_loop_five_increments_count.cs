// vybe-test: csharp/csharp_interlocked_atomic/interlocked_loop_five_increments_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int counter = 0;
for (int i = 0; i < 5; i++) System.Threading.Interlocked.Increment(ref counter);
Console.WriteLine(counter);
