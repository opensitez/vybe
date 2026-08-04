// vybe-test: csharp/csharp_datetime_offset/datetime_to_universal_time_converts_to_utc_kind
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

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

var local=new System.DateTime(2024,1,15,12,0,0,System.DateTimeKind.Local);
var utc=local.ToUniversalTime();
__P((utc.Kind).ToString());
__Check("Utc");
