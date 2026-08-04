// vybe-test: csharp/csharp_events_advanced/eventhandler_passes_sender_and_args_values
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using System; class MessageEventArgs : EventArgs { public string Text { get; set; } } class Channel { public event EventHandler<MessageEventArgs> Sent; public void Emit(string text) { Sent(this, new MessageEventArgs { Text = text }); } } var channel = new Channel(); channel.Sent += (sender, args) => __P((args.Text).ToString()); channel.Emit("hello");
__Check("hello");
