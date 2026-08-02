// vybe-test: csharp/csharp_event_advanced/multicast_event_all_subscribers_called_in_order
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{public event System.Action Click;}
var b=new Btn();
var log=new System.Collections.Generic.List<string>();
b.Click+=()=>log.Add("a");
b.Click+=()=>log.Add("b");
b.Click?.Invoke();
__Check((string.Join(",",log)).ToString(), "a,b");
