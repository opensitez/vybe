// vybe-test: csharp/csharp_custom_event_accessors/custom_event_subscriber_count_property
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Subscribers=>_c==null?0:_c.GetInvocationList().Length;} var b=new Btn(); b.Click+=()=>{}; b.Click+=()=>{}; __Check((b.Subscribers).ToString(), "2");
