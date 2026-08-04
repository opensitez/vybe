// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_add_hours_adjusts_wall_time
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

var dto=new System.DateTimeOffset(2024,1,1,20,0,0,System.TimeSpan.Zero);
var next=dto.AddHours(5);
__P((next.Day).ToString()); __P((next.Hour).ToString());
__Check("2\n1");
