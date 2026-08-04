// vybe-test: csharp/csharp_delegates_advanced/predicate_combined_with_and
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

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

System.Predicate<int> positive=x=>x>0;
System.Predicate<int> even=x=>x%2==0;
System.Predicate<int> both=x=>positive(x)&&even(x);
__P((both(4)).ToString()); __P((both(-2)).ToString()); __P((both(3)).ToString());
__Check("True\nFalse\nFalse");
