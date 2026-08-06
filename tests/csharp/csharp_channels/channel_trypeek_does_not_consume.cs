// vybe-test: csharp/csharp_channels/channel_trypeek_does_not_consume
// origin: hand-written, expectations validated against dotnet 10.0.100

using System.Threading.Channels;

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var ch = Channel.CreateUnbounded<int>();
ch.Writer.TryWrite(11);
int p;
ch.Reader.TryPeek(out p);
int v;
ch.Reader.TryRead(out v);
__P(p.ToString() + "," + v.ToString() + "," + ch.Reader.Count.ToString());
__Check("11,11,0");
