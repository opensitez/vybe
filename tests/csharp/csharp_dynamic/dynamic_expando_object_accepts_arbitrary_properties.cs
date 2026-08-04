// vybe-test: csharp/csharp_dynamic/dynamic_expando_object_accepts_arbitrary_properties
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

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

dynamic obj=new System.Dynamic.ExpandoObject();
obj.Name="Alice";
obj.Age=30;
__P((obj.Name).ToString()); __P((obj.Age).ToString());
__Check("Alice\n30");
