// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_setter_expression_updates_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Logger { public string last = ""; public string Last { get => last; set => last = value; } }
var l = new Logger(); l.Last = "ok"; __Check((l.Last).ToString(), "ok");
