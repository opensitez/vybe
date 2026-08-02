// vybe-test: csharp/linq_lambdas/event_pattern
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

class Button {
    public event Action OnClick;
    public void Click() { if (OnClick != null) OnClick(); }
}
var btn = new Button();
btn.OnClick += () => Console.WriteLine("clicked!");
btn.Click();
btn.Click();
