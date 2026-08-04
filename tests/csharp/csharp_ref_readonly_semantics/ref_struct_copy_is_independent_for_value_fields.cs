// vybe-test: csharp/csharp_ref_readonly_semantics/ref_struct_copy_is_independent_for_value_fields
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

ref struct Box{public int Item;} var x=new Box(); x.Item=10; var y=x; y.Item=99; __P((x.Item).ToString());
__Check("10");
