// vybe-test: csharp/csharp_events_advanced/eventhandler_passes_sender_and_args_values
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

var channel = new Channel();
channel.Sent += (sender, args) => __P((args.Text).ToString());
channel.Emit("hello");
__Check("hello");

class MessageEventArgs : EventArgs { public string Text { get; set; } }

class Channel { public event EventHandler<MessageEventArgs> Sent; public void Emit(string text) { Sent(this, new MessageEventArgs { Text = text }); } }

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
