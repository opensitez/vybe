// vybe-test: csharp/csharp_constructor_patterns/parameterless_constructor_required_for_generic_new_constraint
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget{public int Value=7;}
T Make<T>() where T:new()=>new T();
__Check((Make<Widget>().Value).ToString(), "7");
