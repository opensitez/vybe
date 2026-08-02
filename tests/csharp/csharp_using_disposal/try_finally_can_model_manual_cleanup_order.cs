// vybe-test: csharp/csharp_using_disposal/try_finally_can_model_manual_cleanup_order
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try { __Check(("body").ToString(), "body"); } finally { __Check(("cleanup").ToString(), "cleanup"); }
