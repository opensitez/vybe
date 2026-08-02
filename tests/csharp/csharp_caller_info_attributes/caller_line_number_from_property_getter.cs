// vybe-test: csharp/csharp_caller_info_attributes/caller_line_number_from_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    int _v = 1;
    public int Value {
        get {
            Trace.Show();
            return _v;
        }
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => __Check((line).ToString(), "5");
}
__Check((new Box().Value).ToString(), "1");
