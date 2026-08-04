// vybe-test: csharp/csharp_record_struct_deep/record_struct_equals_object_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

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

record struct Key(int Id); object o=new Key(5); __P((new Key(5).Equals(o)).ToString());
__Check("True");
