// vybe-test: csharp/csharp_events_advanced/generic_eventhandler_passes_integer_payload
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class NumberEventArgs : EventArgs { public int Value { get; set; } } class Stream { public event EventHandler<NumberEventArgs> Produced; public void Emit(int value) { Produced(this, new NumberEventArgs { Value = value }); } } var stream = new Stream(); stream.Produced += (sender, args) => __Check((args.Value * 2).ToString(), "12"); stream.Emit(6);
