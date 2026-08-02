// vybe-test: csharp/csharp_delegates/event_basic
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

class Button {
    public event Action Click;
    public void Press() {
        if (Click != null) Click();
    }
}
var btn = new Button();
btn.Click += () => Console.WriteLine("clicked!");
btn.Press();
btn.Press();
