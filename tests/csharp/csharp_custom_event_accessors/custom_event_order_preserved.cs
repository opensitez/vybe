// vybe-test: csharp/csharp_custom_event_accessors/custom_event_order_preserved
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} var log=new System.Collections.Generic.List<string>(); var b=new Btn(); b.Click+=()=>log.Add("1"); b.Click+=()=>log.Add("2"); b.Raise(); __Check((string.Join(",",log)).ToString(), "1,2");
