// vybe-test: csharp/csharp_constructor_patterns/base_constructor_called_before_derived_body
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

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

class A{public int Order;public A(){Order=1;}}
class B:A{public int Extra;public B():base(){Extra=2;}}
var b=new B();
__P((b.Order).ToString()); __P((b.Extra).ToString());
__Check("1\n2");
