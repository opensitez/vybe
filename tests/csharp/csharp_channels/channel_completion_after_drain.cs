// vybe-test: csharp/csharp_channels/channel_completion_after_drain
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateUnbounded<int>();
ch.Writer.TryWrite(1);
ch.Writer.Complete();
bool before = ch.Reader.Completion.IsCompleted;
int v;
ch.Reader.TryRead(out v);
bool after = ch.Reader.Completion.IsCompleted;
__P((before ? "y" : "n") + "," + (after ? "y" : "n"));
__Check("n,y");

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
