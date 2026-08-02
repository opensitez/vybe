// vybe-test: csharp/csharp_custom_event_accessors/custom_event_action_int
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Src{System.Action<int> _h; public event System.Action<int> Value{add{_h+=value;} remove{_h-=value;}} public void Set(int v){_h?.Invoke(v);}} int got=0; var s=new Src(); s.Value+=v=>got=v; s.Set(15); __Check((got).ToString(), "15");
