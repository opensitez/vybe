// vybe-test: csharp/csharp_async_advanced/configure_await_false_does_not_resume_on_original_context
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Compute(){
    await System.Threading.Tasks.Task.Delay(1).ConfigureAwait(false);
    return 42;
}
__Check((await Compute()).ToString(), "42");
