// vybe-test: csharp/csharp_pattern_property/is_property_pattern_with_guard_on_captured_var
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair { public int A; public int B; } object o=new Pair{A=4,B=9}; if(o is Pair{A:var a,B:var b} && a<b) __Check((b-a).ToString(), "5");
