// vybe-test: csharp/csharp_pattern_property/switch_expression_property_pattern_capture_amount
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Wallet { public int Balance; } int Read(object o)=>o switch{Wallet{Balance:var b}=>b,_=>-1}; __Check((Read(new Wallet{Balance=42})).ToString(), "42");
