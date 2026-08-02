// vybe-test: csharp/csharp_using_disposal/lock_statement_can_mutate_shared_local_state
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object(); int count = 0; lock (gate) { count += 3; } __Check((count).ToString(), "3");
