// vybe-test: csharp/csharp_interlocked_atomic/interlocked_loop_add_accumulates
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int total = 0;
for (int i = 1; i <= 4; i++) System.Threading.Interlocked.Add(ref total, i);
Console.WriteLine(total);
