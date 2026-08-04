// vybe-test: csharp/csharp_generic_methods/generic_action_parameterised_with_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

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

void ForEach<T>(T[] items,System.Action<T> action){
    foreach(var i in items) action(i);
}
ForEach(new[]{1,2,3},n=>__P((n).ToString()));
__Check("1\n2\n3");
