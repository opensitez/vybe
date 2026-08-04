// vybe-test: csharp/csharp_reflection_activation/get_generic_arguments_reports_type_argument_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

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

class Box<T> { } __P((typeof(Box<int>).GetGenericArguments()[0].Name).ToString());
__Check("Int32");
