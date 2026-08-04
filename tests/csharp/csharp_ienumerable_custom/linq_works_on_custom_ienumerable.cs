// vybe-test: csharp/csharp_ienumerable_custom/linq_works_on_custom_ienumerable
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

class Odds:System.Collections.Generic.IEnumerable<int>{
    int _count;
    public Odds(int count){_count=count;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=0;i<_count;i++) yield return 2*i+1;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}
__P((new Odds(4).Sum()).ToString());
__Check("16");
