// vybe-test: csharp/type_features/string_concat_as_interpolation
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var name = "World";
        var age = 25;
        __Check(("Hello " + name + ", age " + age).ToString(), "Hello World, age 25");
