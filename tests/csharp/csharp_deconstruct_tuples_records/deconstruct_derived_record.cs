// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_derived_record
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

record Animal(string Name); record Dog(string Name,int Age):Animal(Name); var (name,age)=new Dog("Rex",4); __P((name).ToString()); __P((age).ToString());
__Check("Rex\n4");
