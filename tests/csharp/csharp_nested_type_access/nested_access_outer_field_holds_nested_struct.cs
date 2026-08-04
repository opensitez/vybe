// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_field_holds_nested_struct
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Grid{public struct Cell{public int V;} Cell _c; public Grid(){_c.V=6;} public int Read()=>_c.V;} __P((new Grid().Read()).ToString());
__Check("6");
