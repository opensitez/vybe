// vybe-test: csharp/modern_features/record_with_expression
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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
var p1 = new Person("Alice", 30);
var p2 = p1 with { Age = 31 };
__P((p1).ToString());
__P((p2).ToString());
__Check("Person { Name = Alice, Age = 30 }\nPerson { Name = Alice, Age = 31 }");
