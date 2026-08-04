// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_array_field_reference_cannot_be_replaced_but_elements_can
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

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

class Holder {
    public readonly int[] Data = { 1, 2 };
}
var holder = new Holder();
holder.Data[1] = 9;
__P((holder.Data[1]).ToString());
__Check("9");
