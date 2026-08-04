// vybe-test: csharp/csharp_oop_advanced2/object_type_is_common_base_of_all_classes
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

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

object x=42; object y="hi"; object z=new int[]{};
__P((x is object).ToString());
__P((y is object).ToString());
__P((z is object).ToString());
__Check("True\nTrue\nTrue");
