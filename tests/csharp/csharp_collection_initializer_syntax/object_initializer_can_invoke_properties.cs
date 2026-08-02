// vybe-test: csharp/csharp_collection_initializer_syntax/object_initializer_can_invoke_properties
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class User { public string Name { get; set; } }
var user = new User { Name = "Ada" };
__Check((user.Name).ToString(), "Ada");
