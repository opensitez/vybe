// vybe-test: csharp/csharp_channels/channel_bounded_drain_reopens_room
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateBounded<int>(2);
ch.Writer.TryWrite(1);
ch.Writer.TryWrite(2);
bool full = ch.Writer.TryWrite(3);
int v;
ch.Reader.TryRead(out v);
bool room = ch.Writer.TryWrite(3);
__P((full ? "y" : "n") + "," + (room ? "y" : "n") + "," + ch.Reader.Count.ToString());
__Check("n,y,2");

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
