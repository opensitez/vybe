// vybe-test: csharp/csharp_channels/channel_async_wait_to_read_states
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateUnbounded<int>();
ch.Writer.TryWrite(1);
bool a = await ch.Reader.WaitToReadAsync();
ch.Writer.Complete();
int x;
ch.Reader.TryRead(out x);
bool b = await ch.Reader.WaitToReadAsync();
__P((a ? "y" : "n") + "," + (b ? "y" : "n"));
__Check("y,n");

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
