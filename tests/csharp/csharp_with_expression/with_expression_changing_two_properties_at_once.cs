// vybe-test: csharp/csharp_with_expression/with_expression_changing_two_properties_at_once
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

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

record Person(string Name, int Age);
var p = new Person("Ada", 30);
var updated = p with { Name = "Grace", Age = 31 };
__P((updated.Name).ToString());
__P((updated.Age).ToString());
__Check("Grace\n31");
