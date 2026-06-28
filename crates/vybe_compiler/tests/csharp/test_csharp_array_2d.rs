//! Two-dimensional and three-dimensional arrays: declaration, access, iteration.
use super::helpers::run_csharp;

#[test]
fn two_d_array_set_and_get_by_index() {
    assert_eq!(
        run_csharp(
            r#"int[,] m=new int[3,3];
m[1,2]=99;
Console.WriteLine(m[1,2]);"#
        ),
        &["99"]
    );
}

#[test]
fn two_d_array_row_and_column_lengths() {
    assert_eq!(
        run_csharp(
            r#"int[,] m=new int[4,5];
Console.WriteLine(m.GetLength(0)); Console.WriteLine(m.GetLength(1));"#
        ),
        &["4", "5"]
    );
}

#[test]
fn two_d_array_foreach_visits_all_elements() {
    assert_eq!(
        run_csharp(
            r#"int[,] m={{1,2},{3,4}};
int sum=0; foreach(int n in m) sum+=n;
Console.WriteLine(sum);"#
        ),
        &["10"]
    );
}

#[test]
fn three_d_array_dimension_count() {
    assert_eq!(
        run_csharp(
            r#"int[,,] t=new int[2,3,4];
Console.WriteLine(t.Rank);"#
        ),
        &["3"]
    );
}

#[test]
fn two_d_array_length_is_total_element_count() {
    assert_eq!(
        run_csharp(
            r#"int[,] m=new int[3,4];
Console.WriteLine(m.Length);"#
        ),
        &["12"]
    );
}

#[test]
fn two_d_array_initializer_syntax() {
    assert_eq!(
        run_csharp(
            r#"int[,] m={{1,2,3},{4,5,6}};
Console.WriteLine(m[0,2]); Console.WriteLine(m[1,0]);"#
        ),
        &["3", "4"]
    );
}
