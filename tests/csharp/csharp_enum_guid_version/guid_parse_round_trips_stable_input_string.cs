// vybe-test: csharp/csharp_enum_guid_version/guid_parse_round_trips_stable_input_string
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

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

var text = "00112233-4455-6677-8899-aabbccddeeff"; __P((System.Guid.Parse(text).ToString()).ToString());
__Check("00112233-4455-6677-8899-aabbccddeeff");
