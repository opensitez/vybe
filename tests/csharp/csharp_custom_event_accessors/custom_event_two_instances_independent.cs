// vybe-test: csharp/csharp_custom_event_accessors/custom_event_two_instances_independent
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int a=0,b=0; var x=new Btn(); var y=new Btn(); x.Click+=()=>a++; y.Click+=()=>b++; x.Raise(); __Check((a).ToString(), "1"); __Check((b).ToString(), "0");
