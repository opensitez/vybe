//! `ref`, `out`, and `in` parameter modifiers.
use super::helpers::run_csharp;

#[test]
fn ref_parameter_mutates_caller_variable() {
    assert_eq!(
        run_csharp(
            r#"void Double(ref int x){x*=2;}
int n=5; Double(ref n); Console.WriteLine(n);"#
        ),
        &["10"]
    );
}

#[test]
fn out_parameter_initialised_inside_callee() {
    assert_eq!(
        run_csharp(
            r#"void Minmax(int[] a, out int min, out int max){
    min=a[0]; max=a[0];
    foreach(var v in a){if(v<min)min=v; if(v>max)max=v;}
}
Minmax(new[]{3,1,4,1,5,9}, out int lo, out int hi);
Console.WriteLine(lo); Console.WriteLine(hi);"#
        ),
        &["1", "9"]
    );
}

#[test]
fn out_inline_declaration_in_method_call() {
    assert_eq!(
        run_csharp(
            r#"bool ok = int.TryParse("42", out int result);
Console.WriteLine(ok); Console.WriteLine(result);"#
        ),
        &["True", "42"]
    );
}

#[test]
fn in_parameter_prevents_copy_and_is_readonly() {
    assert_eq!(
        run_csharp(
            r#"int Sum3(in int a, in int b, in int c) => a+b+c;
Console.WriteLine(Sum3(1,2,3));"#
        ),
        &["6"]
    );
}

#[test]
fn ref_local_aliases_array_element() {
    assert_eq!(
        run_csharp(
            r#"int[] arr={1,2,3};
ref int second=ref arr[1];
second=99;
Console.WriteLine(arr[1]);"#
        ),
        &["99"]
    );
}

#[test]
fn ref_return_allows_external_mutation_of_field() {
    assert_eq!(
        run_csharp(
            r#"class Grid{
    int[] data={0,0,0};
    public ref int Cell(int i)=>ref data[i];
    public int Get(int i)=>data[i];
}
var g=new Grid();
g.Cell(1)=7;
Console.WriteLine(g.Get(1));"#
        ),
        &["7"]
    );
}

#[test]
fn multiple_out_parameters_assign_multiple_return_values() {
    assert_eq!(
        run_csharp(
            r#"void Split(string s, out string head, out string tail){
    int mid=s.Length/2;
    head=s.Substring(0,mid); tail=s.Substring(mid);
}
Split("abcdef",out string h,out string t);
Console.WriteLine(h); Console.WriteLine(t);"#
        ),
        &["abc", "def"]
    );
}
