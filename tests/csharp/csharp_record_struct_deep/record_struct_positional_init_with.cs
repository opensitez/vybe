// vybe-test: csharp/csharp_record_struct_deep/record_struct_positional_init_with
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

record struct User(string Name){public int Age{get;init;}} var u=new User("Ada"){Age=30}; var v=u with{Age=31}; __P((u.Age).ToString()); __P((v.Age).ToString());
__Check("30\n31");
