// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_in_loop_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int slot = 0;
for (int i = 1; i <= 3; i++) {
    System.Threading.Interlocked.Exchange(ref slot, i);
}
Console.WriteLine(slot);
