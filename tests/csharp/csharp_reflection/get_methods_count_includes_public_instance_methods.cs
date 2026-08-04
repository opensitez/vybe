// vybe-test: csharp/csharp_reflection/get_methods_count_includes_public_instance_methods
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

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

class Calc { public int Add(int a, int b) => a+b; public int Sub(int a, int b) => a-b; }
__P((typeof(Calc).GetMethods(
    System.Reflection.BindingFlags.Public|System.Reflection.BindingFlags.Instance|
    System.Reflection.BindingFlags.DeclaredOnly).Length).ToString());
__Check("2");
