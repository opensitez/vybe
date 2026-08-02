// vybe-test: csharp/csharp_static_type_behaviors/nested_static_class_exposes_namespaced_helper_method
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class TextTools {
    public static class Parts {
        public static string Join(string a, string b) { return a + "/" + b; }
    }
}
__Check((TextTools.Parts.Join("a", "b")).ToString(), "a/b");
