// vybe-test: csharp/csharp_events_advanced/eventhandler_passes_sender_and_args_values
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class MessageEventArgs : EventArgs { public string Text { get; set; } } class Channel { public event EventHandler<MessageEventArgs> Sent; public void Emit(string text) { Sent(this, new MessageEventArgs { Text = text }); } } var channel = new Channel(); channel.Sent += (sender, args) => __Check((args.Text).ToString(), "hello"); channel.Emit("hello");
