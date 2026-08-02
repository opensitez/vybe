// vybe-test: csharp/csharp_custom_event_accessors/custom_event_count_tracked
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _e; int _count; public event System.Action Tick{add{_e+=value;_count++;} remove{_e-=value;_count--;}} public int Count=>_count; public void Fire(){_e?.Invoke();}} var b=new Btn(); System.Action h=()=>{}; b.Tick+=h; b.Tick+=()=>{}; b.Tick-=h; __Check((b.Count).ToString(), "1");
