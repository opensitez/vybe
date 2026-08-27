// vybe-test: csharp/csharp_channels/channel_trypeek_does_not_consume
// origin: hand-written, expectations validated against dotnet 10.0.100

using static __Harness;
using System.Threading.Channels;

var ch = Channel.CreateUnbounded<int>();
ch.Writer.TryWrite(11);
int p;
ch.Reader.TryPeek(out p);
int v;
ch.Reader.TryRead(out v);
__P(p.ToString() + "," + v.ToString() + "," + ch.Reader.Count.ToString());
__Check("11,11,0");

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
