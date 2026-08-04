// vybe-test: csharp/csharp_events_advanced/generic_eventhandler_passes_integer_payload
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

using System; class NumberEventArgs : EventArgs { public int Value { get; set; } } class Stream { public event EventHandler<NumberEventArgs> Produced; public void Emit(int value) { Produced(this, new NumberEventArgs { Value = value }); } } var stream = new Stream(); stream.Produced += (sender, args) => __P((args.Value * 2).ToString()); stream.Emit(6);
__Check("12");
