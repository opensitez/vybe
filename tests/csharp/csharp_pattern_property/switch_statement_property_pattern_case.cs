// vybe-test: csharp/csharp_pattern_property/switch_statement_property_pattern_case
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

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

class Node { public int Id; } object o=new Node{Id=5}; string tag=""; switch(o){case Node{Id:5}:tag="match";break;default:tag="miss";break;} __P((tag).ToString());
__Check("match");
