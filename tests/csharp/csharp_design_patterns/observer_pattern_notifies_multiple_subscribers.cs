// vybe-test: csharp/csharp_design_patterns/observer_pattern_notifies_multiple_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Button{public event System.Action Clicked;}
var b=new Button();
int a=0,c=0;
b.Clicked+=()=>a++;
b.Clicked+=()=>c++;
b.Clicked?.Invoke();
__Check((a).ToString(), "1"); __Check((c).ToString(), "1");
