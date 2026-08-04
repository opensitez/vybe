// vybe-test: csharp/csharp_yield_iterators_core/yield_break_inside_nested_loop
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

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

System.Collections.Generic.IEnumerable<int> Grid(int rows,int cols){for(int r=0;r<rows;r++){for(int c=0;c<cols;c++){if(r==1&&c==1)yield break;yield return r*10+c;}}}
__P((string.Join(",",Grid(3,3))).ToString());
__Check("0,1,2,10,11");
