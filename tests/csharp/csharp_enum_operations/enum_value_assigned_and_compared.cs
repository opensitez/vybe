// vybe-test: csharp/csharp_enum_operations/enum_value_assigned_and_compared
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color{Red,Green,Blue} var c=Color.Green; __Check((c==Color.Green).ToString(), "True");
