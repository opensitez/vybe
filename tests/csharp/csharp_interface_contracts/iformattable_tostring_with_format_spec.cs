// vybe-test: csharp/csharp_interface_contracts/iformattable_tostring_with_format_spec
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

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

System.IFormattable value = (object)3.14;
__P((value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture)).ToString());
__Check("3.1");
