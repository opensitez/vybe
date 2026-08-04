// vybe-test: csharp/csharp_event_advanced/multicast_event_all_subscribers_called_in_order
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{public event System.Action Click;}
var b=new Btn();
var log=new System.Collections.Generic.List<string>();
b.Click+=()=>log.Add("a");
b.Click+=()=>log.Add("b");
b.Click?.Invoke();
__P((string.Join(",",log)).ToString());
__Check("a,b");
