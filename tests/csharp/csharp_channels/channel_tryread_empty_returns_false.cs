// vybe-test: csharp/csharp_channels/channel_tryread_empty_returns_false
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateUnbounded<int>();
int v;
__P(ch.Reader.TryRead(out v) ? "y" : "n");
__Check("n");

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
