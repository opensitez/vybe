// vybe-test: csharp/csharp_custom_event_accessors/custom_event_count_zero_no_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Count=>_c==null?0:_c.GetInvocationList().Length; public void Raise(){_c?.Invoke();}} var b=new Btn(); b.Raise(); __Check((b.Count).ToString(), "0");
