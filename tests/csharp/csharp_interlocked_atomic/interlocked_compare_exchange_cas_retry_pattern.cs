// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_cas_retry_pattern
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int slot = 0;
int observed = slot;
int desired = observed + 1;
while (observed != System.Threading.Interlocked.CompareExchange(ref slot, desired, observed)) {
    observed = slot;
    desired = observed + 1;
}
Console.WriteLine(slot);
