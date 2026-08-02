// vybe-test: csharp/csharp_enum_operations/enum_parse_converts_string_name_to_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Day{Mon,Tue,Wed,Thu,Fri}
var d = (Day)System.Enum.Parse(typeof(Day),"Wed");
__Check((d).ToString(), "Wed");
