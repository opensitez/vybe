// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_struct_element_type
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

struct Pt{public int X;} System.Collections.Generic.IEnumerable<Pt> Points(){yield return new Pt{X=1};yield return new Pt{X=2};}
__P((Points().Sum(p=>p.X)).ToString());
__Check("3");
