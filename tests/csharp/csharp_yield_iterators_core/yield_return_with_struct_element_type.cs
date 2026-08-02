// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_struct_element_type
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Pt{public int X;} System.Collections.Generic.IEnumerable<Pt> Points(){yield return new Pt{X=1};yield return new Pt{X=2};}
__Check((Points().Sum(p=>p.X)).ToString(), "3");
