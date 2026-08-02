// vybe-test: csharp/csharp_local_function_static/local_function_captures_class_field
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box{public int Value=5; int Scale(int n){int S(int x)=>x*Value; return S(n);}} var b=new Box(); __Check((b.Scale(3)).ToString(), "15");
