// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_nested_record_field
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

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

// `var (n) = …` is not C# — a deconstruction declaration needs two or more
// targets (dotnet: CS1001 on the original form). Rewritten to a two-field
// record so the nested-field read the test is named for still happens;
// dotnet prints GOT[9|1] for this exact program.
record Inner(int N); record Outer(Inner I, int Tag); var (i, t) = new Outer(new Inner(9), 1); __P((i.N).ToString()); __P((t).ToString());
__Check("9\n1");
