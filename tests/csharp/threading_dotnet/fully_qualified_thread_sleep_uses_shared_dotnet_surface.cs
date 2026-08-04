// vybe-test: csharp/threading_dotnet/fully_qualified_thread_sleep_uses_shared_dotnet_surface
// origin: languages/csharp/tests/csharp/test_threading_dotnet.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

__P(("before").ToString());
        System.Threading.Thread.Sleep(1);
        __P(("after").ToString());
__Check("before\nafter");
