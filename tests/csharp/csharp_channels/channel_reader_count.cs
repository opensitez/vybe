// vybe-test: csharp/csharp_channels/channel_reader_count
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
ch.Writer.TryWrite(2);
ch.Writer.TryWrite(3);
__P(ch.Reader.Count.ToString());
__Check("3");
