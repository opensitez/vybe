// vybe-test: csharp/csharp_generics_constraints/generic_method_with_where_t_base_class_works_on_derived_input
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

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

class Person { public string Name = "Ada"; } class Admin : Person { } string Read<T>(T person) where T : Person { return person.Name; } __P((Read(new Admin())).ToString());
__Check("Ada");
