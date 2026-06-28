//! Advanced array operations: multi-dimensional, jagged initialisation, `Array` static methods.
use super::helpers::run_csharp;

#[test]
fn multidimensional_array_length_at_each_dimension() {
    assert_eq!(
        run_csharp(
            r#"int[,] grid=new int[3,4];
Console.WriteLine(grid.GetLength(0)); Console.WriteLine(grid.GetLength(1));"#
        ),
        &["3", "4"]
    );
}

#[test]
fn jagged_array_inner_arrays_have_independent_lengths() {
    assert_eq!(
        run_csharp(
            r#"int[][] j=new int[3][];
j[0]=new int[1]; j[1]=new int[2]; j[2]=new int[3];
Console.WriteLine(j[0].Length); Console.WriteLine(j[2].Length);"#
        ),
        &["1", "3"]
    );
}

#[test]
fn array_sort_then_binary_search_finds_element() {
    assert_eq!(
        run_csharp(
            r#"int[] arr={5,3,1,4,2};
System.Array.Sort(arr);
int idx=System.Array.BinarySearch(arr,4);
Console.WriteLine(idx);"#
        ),
        &["3"]
    );
}

#[test]
fn array_fill_sets_all_elements_to_value() {
    assert_eq!(
        run_csharp(
            r#"int[] arr=new int[5];
System.Array.Fill(arr,7);
Console.WriteLine(arr[2]);"#
        ),
        &["7"]
    );
}

#[test]
fn array_convert_all_transforms_each_element() {
    assert_eq!(
        run_csharp(
            r#"int[] src={1,2,3};
string[] dst=System.Array.ConvertAll(src,n=>n.ToString()+"x");
Console.WriteLine(dst[1]);"#
        ),
        &["2x"]
    );
}

#[test]
fn array_true_for_all_validates_all_elements() {
    assert_eq!(
        run_csharp(
            r#"int[] arr={2,4,6,8};
Console.WriteLine(System.Array.TrueForAll(arr,n=>n%2==0));"#
        ),
        &["True"]
    );
}

#[test]
fn array_create_instance_via_reflection_type() {
    assert_eq!(
        run_csharp(
            r#"var arr=(int[])System.Array.CreateInstance(typeof(int),5);
arr[3]=99;
Console.WriteLine(arr[3]);"#
        ),
        &["99"]
    );
}
