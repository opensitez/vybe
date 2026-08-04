// vybe-test: csharp/csharp_generic_methods/generic_method_swap_exchanges_two_values_via_ref
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

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

void Swap<T>(ref T a,ref T b){T tmp=a;a=b;b=tmp;}
int x=1,y=2; Swap(ref x,ref y);
__P((x).ToString()); __P((y).ToString());
__Check("2\n1");
