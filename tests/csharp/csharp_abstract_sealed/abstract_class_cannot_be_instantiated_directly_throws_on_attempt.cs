// vybe-test: csharp/csharp_abstract_sealed/abstract_class_cannot_be_instantiated_directly_throws_on_attempt
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Base { }
string result = "ok";
try {
    var obj = System.Activator.CreateInstance(typeof(Base));
    result = "created";
} catch (System.MemberAccessException) {
    result = "blocked";
} catch (System.Exception) {
    result = "blocked";
}
__Check((result).ToString(), "blocked");
