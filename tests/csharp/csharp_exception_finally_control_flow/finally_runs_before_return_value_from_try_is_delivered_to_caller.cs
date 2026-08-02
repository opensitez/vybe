// vybe-test: csharp/csharp_exception_finally_control_flow/finally_runs_before_return_value_from_try_is_delivered_to_caller
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Pick() {
    try {
        return 2;
    } finally {
        __Check(("cleanup").ToString(), "cleanup");
    }
}
__Check((Pick()).ToString(), "2");
