// vybe-test: csharp/threading_dotnet/fully_qualified_thread_sleep_uses_shared_dotnet_surface
// origin: languages/csharp/tests/csharp/test_threading_dotnet.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("before").ToString(), "before");
        System.Threading.Thread.Sleep(1);
        __Check(("after").ToString(), "after");
