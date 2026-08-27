// vybe-test: csharp/csharp_channels/channel_complete_drains_then_false
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateUnbounded<int>();
ch.Writer.TryWrite(7);
ch.Writer.Complete();
int a;
bool oka = ch.Reader.TryRead(out a);
int b;
bool okb = ch.Reader.TryRead(out b);
__P((oka ? "y" : "n") + "," + a.ToString() + "," + (okb ? "y" : "n"));
__Check("y,7,n");

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
