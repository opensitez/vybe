// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_can_access_properties_of_extended_type
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int Width, Height; }
static class BoxExt { public static int Area(this Box b) => b.Width*b.Height; }
__P((new Box{Width=3,Height=4}.Area()).ToString());
__Check("12");
