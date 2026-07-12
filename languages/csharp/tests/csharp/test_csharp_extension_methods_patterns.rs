//! Extension method declaration, chaining, and use on built-in types.
use super::helpers::run_csharp;

#[test]
fn extension_method_on_int_adds_new_behaviour() {
    assert_eq!(
        run_csharp(
            r#"static class IntExt { public static bool IsEven(this int n) => n%2==0; }
Console.WriteLine(4.IsEven()); Console.WriteLine(3.IsEven());"#
        ),
        &["True", "False"]
    );
}

#[test]
fn extension_method_on_string_chains_after_built_in_method() {
    assert_eq!(
        run_csharp(
            r#"static class StrExt { public static string Shout(this string s) => s.ToUpper() + "!"; }
Console.WriteLine("hello".Shout());"#
        ),
        &["HELLO!"]
    );
}

#[test]
fn extension_method_on_ienumerable_implements_custom_linq_step() {
    assert_eq!(
        run_csharp(
            r#"static class SeqExt {
    public static System.Collections.Generic.IEnumerable<T> Every<T>(
        this System.Collections.Generic.IEnumerable<T> src, int n) {
        int i=0; foreach(var x in src) { if(i++%n==0) yield return x; }
    }
}
var result = new[]{1,2,3,4,5,6}.Every(2);
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["1", "3", "5"]
    );
}

#[test]
fn extension_method_can_access_properties_of_extended_type() {
    assert_eq!(
        run_csharp(
            r#"class Box { public int Width, Height; }
static class BoxExt { public static int Area(this Box b) => b.Width*b.Height; }
Console.WriteLine(new Box{Width=3,Height=4}.Area());"#
        ),
        &["12"]
    );
}

#[test]
fn extension_method_on_nullable_int_handles_null_gracefully() {
    assert_eq!(
        run_csharp(
            r#"static class NullableExt { public static int OrZero(this int? n) => n ?? 0; }
int? present=5, absent=null;
Console.WriteLine(present.OrZero()); Console.WriteLine(absent.OrZero());"#
        ),
        &["5", "0"]
    );
}
