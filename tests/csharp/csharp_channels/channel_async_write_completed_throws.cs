// vybe-test: csharp/csharp_channels/channel_async_write_completed_throws
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
await ch.Writer.WriteAsync(1);
ch.Writer.Complete();
try {
    await ch.Writer.WriteAsync(2);
    __P("no-throw");
} catch (Exception) {
    __P("closed");
}
// Completion closes the WRITE side only — the backlog stays readable.
int v = await ch.Reader.ReadAsync();
__P("read " + v);
__Check("closed\nread 1");
