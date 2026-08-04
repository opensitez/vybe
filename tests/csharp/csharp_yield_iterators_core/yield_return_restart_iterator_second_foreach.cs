// vybe-test: csharp/csharp_yield_iterators_core/yield_return_restart_iterator_second_foreach
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

System.Collections.Generic.IEnumerable<int> Two(){yield return 10;yield return 20;}
int sum=0; foreach(var n in Two()) sum+=n; foreach(var n in Two()) sum+=n; __P((sum).ToString());
__Check("60");
