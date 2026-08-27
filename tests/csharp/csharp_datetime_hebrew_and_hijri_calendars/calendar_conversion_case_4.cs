// vybe-test: csharp/csharp_datetime_hebrew_and_hijri_calendars/calendar_conversion_case_4

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var cal = new System.Globalization.HijriCalendar();
var dt = new DateTime(2026, 1, 4);
int year = cal.GetYear(dt);
__P((year > 1400).ToString());
__Check("True");
