// vybe-test: csharp/csharp_reflection_emit/field_info_get_and_set_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

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

class Box{public int V;}
var fi=typeof(Box).GetField("V");
var obj=new Box();
fi.SetValue(obj,55);
__P((fi.GetValue(obj)).ToString());
__Check("55");
