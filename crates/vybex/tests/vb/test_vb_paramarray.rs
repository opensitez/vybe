use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_paramarray(values: &[i32]) {
    let joined = values.iter().map(|value| value.to_string()).collect::<Vec<_>>().join(", ");
    let total: i32 = values.iter().sum();
    let src = format!(r#"
Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll({joined}))
    End Sub
End Module
"#, joined = joined);
    assert_vb_output_owned(src, vec![total.to_string()]);
}

macro_rules! paramarray_cases {
    ($($name:ident => [$($value:expr),* $(,)?]),* $(,)?) => {
        $(#[test] fn $name() { assert_paramarray(&[$($value),*]); })*
    };
}

paramarray_cases! {
    paramarray_001 => [1],
    paramarray_002 => [1, 2],
    paramarray_003 => [1, 2, 3],
    paramarray_004 => [2, 4, 6],
    paramarray_005 => [3, 6, 9, 12],
    paramarray_006 => [5, 5, 5],
    paramarray_007 => [8, 1, 1],
    paramarray_008 => [10, 20],
    paramarray_009 => [7, 3, 2, 1],
    paramarray_010 => [9, 9, 9, 9],
    paramarray_011 => [4, 8, 12, 16],
    paramarray_012 => [11, 1],
    paramarray_013 => [13, 2, 1],
    paramarray_014 => [14, 7, 0],
    paramarray_015 => [15, 5, 5, 5],
    paramarray_016 => [16, 4],
    paramarray_017 => [18, 2, 2, 2],
    paramarray_018 => [20, 1, 1, 1, 1],
    paramarray_019 => [21, 3, 6],
    paramarray_020 => [24, 8, 4],
    paramarray_021 => [25, 25],
    paramarray_022 => [27, 1, 2, 3],
    paramarray_023 => [28, 7, 7],
    paramarray_024 => [30, 5, 10],
    paramarray_025 => [32, 4, 4, 4],
    paramarray_026 => [33, 11, 22],
    paramarray_027 => [35, 5, 15],
    paramarray_028 => [36, 6, 6, 6],
    paramarray_029 => [40, 10],
    paramarray_030 => [42, 1, 2, 3, 4],
    paramarray_031 => [45, 5],
    paramarray_032 => [48, 2, 2],
    paramarray_033 => [50, 10, 5],
    paramarray_034 => [54, 3, 3, 3],
    paramarray_035 => [56, 7, 8],
    paramarray_036 => [60, 6, 6],
    paramarray_037 => [63, 9],
    paramarray_038 => [70, 7, 7, 7],
    paramarray_039 => [72, 8, 8],
    paramarray_040 => [81, 9, 9],
}