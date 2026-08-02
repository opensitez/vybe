// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_class_equality
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Tag { public string Name; public static bool operator ==(Tag a, Tag b) => a.Name == b.Name; public static bool operator !=(Tag a, Tag b) => !(a == b); }
__Check((new Tag { Name = "x" } == new Tag { Name = "x" }).ToString(), "True"); __Check((new Tag { Name = "a" } != new Tag { Name = "b" }).ToString(), "True");
