// vybe-test: csharp/csharp_yield_advanced/yield_return_with_complex_state_machine
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

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

System.Collections.Generic.IEnumerable<string> Words(string s){
    var parts=s.Split(' ');
    foreach(var p in parts) if(p.Length>0) yield return p;
}
__P((string.Join("|",Words("hello  world  foo"))).ToString());
__Check("hello|world|foo");
