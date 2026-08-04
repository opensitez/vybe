// vybe-test: csharp/csharp_collection_initializer_syntax/object_initializer_can_invoke_properties
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

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

class User { public string Name { get; set; } }
var user = new User { Name = "Ada" };
__P((user.Name).ToString());
__Check("Ada");
