// vybe-test: csharp/csharp_nested_classes/nested_static_class_provides_utility_methods
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Parser{
    public static class Helpers{public static int ToInt(string s)=>int.Parse(s);}
}
__Check((Parser.Helpers.ToInt("99")).ToString(), "99");
