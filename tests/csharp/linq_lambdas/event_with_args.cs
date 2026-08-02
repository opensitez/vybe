// vybe-test: csharp/linq_lambdas/event_with_args
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

class Timer {
    public event Action<int> OnTick;
    public void Tick(int count) { if (OnTick != null) OnTick(count); }
}
var t = new Timer();
t.OnTick += n => Console.WriteLine("tick " + n);
t.Tick(1);
t.Tick(2);
