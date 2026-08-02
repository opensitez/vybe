// vybe-test: csharp/csharp_events/event_fires_single_subscriber
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Button {
    public event System.EventHandler Clicked;
    public void Click() => Clicked?.Invoke(this, System.EventArgs.Empty);
}
int count = 0;
var btn = new Button();
btn.Clicked += (s, e) => count++;
btn.Click();
__Check((count).ToString(), "1");
