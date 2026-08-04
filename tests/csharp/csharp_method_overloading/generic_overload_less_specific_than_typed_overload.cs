// vybe-test: csharp/csharp_method_overloading/generic_overload_less_specific_than_typed_overload
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

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

string Foo<T>(T v)=>"generic";
string Foo(int v)=>"specific";
__P((Foo(1)).ToString());
__P((Foo("x")).ToString());
__Check("specific\ngeneric");
