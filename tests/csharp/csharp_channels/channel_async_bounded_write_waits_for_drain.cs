// vybe-test: csharp/csharp_channels/channel_async_bounded_write_waits_for_drain
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateBounded<int>(1);
await ch.Writer.WriteAsync(1);
var t = Task.Run(() => {
    for (int i = 0; i < 2000; i++) { }
    int x;
    ch.Reader.TryRead(out x);
});
await ch.Writer.WriteAsync(2);
__P(ch.Reader.Count.ToString());
__Check("1");

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
