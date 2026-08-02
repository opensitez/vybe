// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_string_normalization
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface INorm<T> where T:INorm<T>{static abstract T Normalize(string s);}
struct Text:INorm<Text>{public string Value; public static Text Normalize(string s)=>new Text{Value=s.Trim().ToLower()};}
__Check((Text.Normalize(" Ab ").Value).ToString(), "ab");
