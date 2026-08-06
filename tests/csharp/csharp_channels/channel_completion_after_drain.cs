// vybe-test: csharp/csharp_channels/channel_completion_after_drain
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
ch.Writer.TryWrite(1);
ch.Writer.Complete();
bool before = ch.Reader.Completion.IsCompleted;
int v;
ch.Reader.TryRead(out v);
bool after = ch.Reader.Completion.IsCompleted;
__P((before ? "y" : "n") + "," + (after ? "y" : "n"));
__Check("n,y");
