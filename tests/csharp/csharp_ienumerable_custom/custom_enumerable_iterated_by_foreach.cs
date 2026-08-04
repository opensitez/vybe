// vybe-test: csharp/csharp_ienumerable_custom/custom_enumerable_iterated_by_foreach
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

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

class UpTo:System.Collections.Generic.IEnumerable<int>{
    int _max;
    public UpTo(int max){_max=max;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=1;i<=_max;i++) yield return i;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}
int sum=0; foreach(var n in new UpTo(5)) sum+=n;
__P((sum).ToString());
__Check("15");
