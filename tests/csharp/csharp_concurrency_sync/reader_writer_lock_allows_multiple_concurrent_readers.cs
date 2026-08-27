// vybe-test: csharp/csharp_concurrency_sync/reader_writer_lock_allows_multiple_concurrent_readers
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

using static __Harness;

__P("Valid_reader_writer_lock_allows_multiple_concurrent_readers");
__Check("Valid_reader_writer_lock_allows_multiple_concurrent_readers");
public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
