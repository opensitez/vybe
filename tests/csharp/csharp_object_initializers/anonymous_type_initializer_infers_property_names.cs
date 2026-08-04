// vybe-test: csharp/csharp_object_initializers/anonymous_type_initializer_infers_property_names
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

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

string name="Alice"; int age=30;
var anon=new{name,age};
__P((anon.name).ToString()); __P((anon.age).ToString());
__Check("Alice\n30");
