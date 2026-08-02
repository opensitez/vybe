// vybe-test: csharp/csharp_using_disposal/multiple_try_finally_blocks_execute_independently
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try { __Check(("one").ToString(), "one"); } finally { __Check(("cleanup-one").ToString(), "cleanup-one"); } try { __Check(("two").ToString(), "two"); } finally { __Check(("cleanup-two").ToString(), "cleanup-two"); }
