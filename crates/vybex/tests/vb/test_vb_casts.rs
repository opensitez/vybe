use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_casts(value: i32) {
    let src = format!(r#"
Module M
    Sub Main()
        Dim boxed As Object = {value}
        Console.WriteLine(DirectCast(boxed, Integer))
        Console.WriteLine(TryCast(boxed, Integer))
    End Sub
End Module
"#, value = value);
    assert_vb_output_owned(src, vec![value.to_string(), value.to_string()]);
}

macro_rules! cast_cases {
    ($($name:ident => $value:expr),* $(,)?) => {
        $(#[test] fn $name() { assert_casts($value); })*
    };
}

cast_cases! {
    casts_001 => 1,
    casts_002 => 2,
    casts_003 => 3,
    casts_004 => 4,
    casts_005 => 5,
    casts_006 => 10,
    casts_007 => 15,
    casts_008 => 20,
    casts_009 => 25,
    casts_010 => 30,
    casts_011 => 35,
    casts_012 => 40,
    casts_013 => 45,
    casts_014 => 50,
    casts_015 => 55,
    casts_016 => 60,
    casts_017 => 65,
    casts_018 => 70,
    casts_019 => 75,
    casts_020 => 80,
}