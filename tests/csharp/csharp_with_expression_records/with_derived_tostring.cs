// vybe-test: csharp/csharp_with_expression_records/with_derived_tostring
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

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

record Animal(string Name); record Cat(string Name,string Color):Animal(Name); var d=(new Cat("M","W")) with{Color="B"}; __P((d.ToString().Contains("B")).ToString());
__Check("True");
