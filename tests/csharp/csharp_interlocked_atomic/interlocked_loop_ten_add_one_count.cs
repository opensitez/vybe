// vybe-test: csharp/csharp_interlocked_atomic/interlocked_loop_ten_add_one_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int counter = 0;
for (int i = 0; i < 10; i++) System.Threading.Interlocked.Add(ref counter, 1);
Console.WriteLine(counter);
