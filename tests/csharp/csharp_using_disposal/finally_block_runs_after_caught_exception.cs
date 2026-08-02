// vybe-test: csharp/csharp_using_disposal/finally_block_runs_after_caught_exception
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try { throw new System.Exception(); } catch (System.Exception) { __Check(("caught").ToString(), "caught"); } finally { __Check(("finally").ToString(), "finally"); }
