// vybe-test: csharp/csharp_with_expression_records/with_nominal_two_inits
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

record Theme{public string Name{get;init;} public int Ver{get;init;}} var u=(new Theme{Name="dark",Ver=1}) with{Name="light",Ver=2}; __P((u.Name).ToString()); __P((u.Ver).ToString());
__Check("light\n2");
