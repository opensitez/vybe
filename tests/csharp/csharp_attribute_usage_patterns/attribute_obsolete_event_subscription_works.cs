// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_event_subscription_works
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Btn{[Obsolete("old")] public event Action Click; public void Fire(){Click?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Fire(); __Check((n).ToString(), "1");
