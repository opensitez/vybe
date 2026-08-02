// vybe-test: csharp/csharp_concurrency_sync/interlocked_compare_exchange_sets_only_when_expected
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int val=0;
int original=System.Threading.Interlocked.CompareExchange(ref val,99,0);
__Check((original).ToString(), "0"); __Check((val).ToString(), "99");
