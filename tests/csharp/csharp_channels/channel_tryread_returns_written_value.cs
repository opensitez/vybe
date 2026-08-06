// vybe-test: csharp/csharp_channels/channel_tryread_returns_written_value
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
ch.Writer.TryWrite(42);
int v;
bool ok = ch.Reader.TryRead(out v);
__P((ok ? "y" : "n") + "," + v.ToString());
__Check("y,42");
