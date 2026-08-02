// vybe-test: csharp/csharp_events/multiple_subscribers_all_called_in_order
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Emitter {
    public event System.Action<string> Signal;
    public void Emit(string v) => Signal?.Invoke(v);
}
string log = "";
var e = new Emitter();
e.Signal += v => log += "A";
e.Signal += v => log += "B";
e.Emit("x");
__Check((log).ToString(), "AB");
