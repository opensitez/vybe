// vybe-test: csharp/csharp_events/custom_event_args_carries_data_to_handler
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class DataArgs : System.EventArgs { public int Value; }
class Source {
    public event System.EventHandler<DataArgs> Changed;
    public void Change(int v) => Changed?.Invoke(this, new DataArgs{Value=v});
}
int received = 0;
var src = new Source();
src.Changed += (s, e) => received = e.Value;
src.Change(77);
__Check((received).ToString(), "77");
