// vybe-test: csharp/csharp_null_propagation/null_conditional_property_access_returns_value_for_non_null_object
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class User { public string Name { get; set; } } var user = new User { Name = "Ada" }; __Check((user?.Name).ToString(), "Ada");
