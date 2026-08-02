// vybe-test: csharp/csharp_anonymous_types/two_anonymous_types_with_same_shape_are_equal
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new{X=1,Y=2}; var b=new{X=1,Y=2};
__Check((a.Equals(b)).ToString(), "True");
