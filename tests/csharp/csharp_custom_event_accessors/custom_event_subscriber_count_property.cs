// vybe-test: csharp/csharp_custom_event_accessors/custom_event_subscriber_count_property
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

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

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Subscribers=>_c==null?0:_c.GetInvocationList().Length;} var b=new Btn(); b.Click+=()=>{}; b.Click+=()=>{}; __P((b.Subscribers).ToString());
__Check("2");
