// vybe-test: csharp/csharp_constructor_chains/struct_constructor_can_delegate_to_field_assignment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

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

struct Pair { public int Left; public int Right; public Pair(int left, int right) { Left = left; Right = right; } } var pair = new Pair(2, 8); __P((pair.Left + pair.Right).ToString());
__Check("10");
