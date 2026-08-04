// vybe-test: csharp/csharp_abstract_sealed/abstract_class_cannot_be_instantiated_directly_throws_on_attempt
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

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

abstract class Base { }
string result = "ok";
try {
    var obj = System.Activator.CreateInstance(typeof(Base));
    result = "created";
} catch (System.MemberAccessException) {
    result = "blocked";
} catch (System.Exception) {
    result = "blocked";
}
__P((result).ToString());
__Check("blocked");
