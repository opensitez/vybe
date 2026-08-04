// vybe-test: csharp/csharp_design_patterns/observer_pattern_notifies_multiple_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

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

class Button{public event System.Action Clicked;}
var b=new Button();
int a=0,c=0;
b.Clicked+=()=>a++;
b.Clicked+=()=>c++;
b.Clicked?.Invoke();
__P((a).ToString()); __P((c).ToString());
__Check("1\n1");
