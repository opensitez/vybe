// vybe-test: csharp/csharp_delegate_types/multicast_delegate_invokes_both_handlers_in_order
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action log = () => __Check(("a").ToString(), "a");
log += () => __Check(("b").ToString(), "b");
log();
