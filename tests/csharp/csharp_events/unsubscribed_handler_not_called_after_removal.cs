// vybe-test: csharp/csharp_events/unsubscribed_handler_not_called_after_removal
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
System.EventHandler h = (s, e) => count++;
var btn = new Button();
btn.Clicked += h;
btn.Clicked -= h;
btn.Click();
__Check((count).ToString(), "0");
