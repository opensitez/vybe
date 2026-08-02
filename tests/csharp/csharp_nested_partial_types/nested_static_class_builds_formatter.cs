// vybe-test: csharp/csharp_nested_partial_types/nested_static_class_builds_formatter
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Report {
    public static class Formatter {
        public static string Line(string key, int value) { return key + ":" + value; }
    }
}
__Check((Report.Formatter.Line("count", 3)).ToString(), "count:3");
