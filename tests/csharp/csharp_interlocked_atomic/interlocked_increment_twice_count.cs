// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_twice_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int counter = 0;
System.Threading.Interlocked.Increment(ref counter);
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
