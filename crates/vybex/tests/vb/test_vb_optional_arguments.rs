use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_optional_arguments(name: &str, default_prefix: &str, explicit_prefix: &str, suffix: &str) {
    let src = format!(r#"
Module M
    Function Decorate(name As String, Optional prefix As String = "{default_prefix}", Optional suffix As String = "{suffix}") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Decorate("{name}"))
        Console.WriteLine(Decorate("{name}", "{explicit_prefix}"))
        Console.WriteLine(Decorate("{name}", "{explicit_prefix}", "done"))
    End Sub
End Module
"#, name = name, default_prefix = default_prefix, explicit_prefix = explicit_prefix, suffix = suffix);
    assert_vb_output_owned(src, vec![
        format!("{}:{}:{}", default_prefix, name, suffix),
        format!("{}:{}:{}", explicit_prefix, name, suffix),
        format!("{}:{}:done", explicit_prefix, name),
    ]);
}

macro_rules! optional_argument_cases {
    ($($name:ident => ($name_value:expr, $default_prefix:expr, $explicit_prefix:expr, $suffix:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_optional_arguments($name_value, $default_prefix, $explicit_prefix, $suffix); })*
    };
}

optional_argument_cases! {
    optional_arguments_001 => ("Ada", "Hello", "Hi", "!"),
    optional_arguments_002 => ("Bob", "Welcome", "Yo", "?"),
    optional_arguments_003 => ("Cora", "Start", "Begin", "."),
    optional_arguments_004 => ("Dax", "Outer", "Inner", "!"),
    optional_arguments_005 => ("Eli", "North", "South", "*"),
    optional_arguments_006 => ("Faye", "Red", "Blue", "!"),
    optional_arguments_007 => ("Gus", "Open", "Close", "?"),
    optional_arguments_008 => ("Hope", "Fast", "Slow", "."),
    optional_arguments_009 => ("Ivy", "Warm", "Cool", "!"),
    optional_arguments_010 => ("Jade", "Early", "Late", "?"),
    optional_arguments_011 => ("Kai", "Alpha", "Beta", "."),
    optional_arguments_012 => ("Lia", "One", "Two", "!"),
    optional_arguments_013 => ("Moe", "Left", "Right", "?"),
    optional_arguments_014 => ("Nia", "Sun", "Moon", "."),
    optional_arguments_015 => ("Omar", "Top", "Base", "!"),
    optional_arguments_016 => ("Pia", "Mint", "Sage", "?"),
    optional_arguments_017 => ("Quin", "Near", "Far", "."),
    optional_arguments_018 => ("Rex", "Bold", "Calm", "!"),
    optional_arguments_019 => ("Sara", "Soft", "Sharp", "?"),
    optional_arguments_020 => ("Taj", "Fresh", "Dry", "."),
    optional_arguments_021 => ("Una", "Bright", "Dim", "!"),
    optional_arguments_022 => ("Vik", "Prime", "Next", "?"),
    optional_arguments_023 => ("Wren", "Stone", "Glass", "."),
    optional_arguments_024 => ("Xena", "Cloud", "Rain", "!"),
    optional_arguments_025 => ("Yara", "Blue", "Gold", "?"),
    optional_arguments_026 => ("Zed", "Low", "High", "."),
    optional_arguments_027 => ("Ari", "Quiet", "Loud", "!"),
    optional_arguments_028 => ("Bea", "First", "Second", "?"),
    optional_arguments_029 => ("Cy", "Wide", "Narrow", "."),
    optional_arguments_030 => ("Dee", "Quick", "Steady", "!"),
    optional_arguments_031 => ("Ena", "Northwest", "Southeast", "?"),
    optional_arguments_032 => ("Fox", "Silver", "Copper", "."),
    optional_arguments_033 => ("Gia", "Lake", "River", "!"),
    optional_arguments_034 => ("Hal", "Root", "Leaf", "?"),
    optional_arguments_035 => ("Ian", "Code", "Data", "."),
    optional_arguments_036 => ("Joy", "Circle", "Square", "!"),
    optional_arguments_037 => ("Kim", "Hot", "Cold", "?"),
    optional_arguments_038 => ("Lou", "Short", "Long", "."),
    optional_arguments_039 => ("Mae", "Solid", "Liquid", "!"),
    optional_arguments_040 => ("Noe", "Plain", "Fancy", "?"),
}