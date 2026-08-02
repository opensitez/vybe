// vybe-test: csharp/csharp_using_disposal/lock_statement_serializes_body_execution
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object(); lock (gate) { __Check(("locked").ToString(), "locked"); }
