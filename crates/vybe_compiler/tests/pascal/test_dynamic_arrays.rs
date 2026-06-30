/// Dynamic arrays (`array of T`): SetLength growth/shrink, assignment, element access.
use super::helpers::run_pascal;

#[test]
fn dynamic_array_setlength_grows_from_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin SetLength(a, 4); WriteLn(Length(a)); WriteLn(High(a)); WriteLn(Low(a)); end."#
        ),
        &["4", "3", "0"]
    );
}

#[test]
fn dynamic_array_setlength_shrink_drops_high_index() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin SetLength(a, 5); a[4] := 99; SetLength(a, 2); WriteLn(Length(a)); WriteLn(High(a)); end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn dynamic_array_setlength_grow_preserves_existing_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin SetLength(a, 2); a[0] := 11; a[1] := 22; SetLength(a, 4); WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["11", "22"]
    );
}

#[test]
fn dynamic_array_literal_constructor_assign() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin a := [7, 8, 9]; WriteLn(a[0]); WriteLn(a[2]); WriteLn(Length(a)); end."#
        ),
        &["7", "9", "3"]
    );
}

#[test]
fn dynamic_array_string_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; var names: array of String; begin names := ['ann', 'bob']; WriteLn(names[1]); WriteLn(Length(names)); end."#
        ),
        &["bob", "2"]
    );
}

#[test]
fn dynamic_array_boolean_flags_count_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var flags: array of Boolean; i, c: Integer; begin flags := [true, false, true]; c := 0; for i := 0 to High(flags) do if flags[i] then c := c + 1; WriteLn(c); end."#
        ),
        &["2"]
    );
}

#[test]
fn dynamic_array_real_values_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var vals: array of Double; s: Double; begin vals := [1.5, 2.5, 3.0]; s := vals[0] + vals[1] + vals[2]; WriteLn(Round(s)); end."#
        ),
        &["7"]
    );
}

#[test]
fn dynamic_array_in_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TB = record Items: array of Integer; end; var b: TB; begin SetLength(b.Items, 2); b.Items[0] := 5; b.Items[1] := 6; WriteLn(b.Items[1]); end."#
        ),
        &["6"]
    );
}

#[test]
fn dynamic_array_returned_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make: array of Integer; begin SetLength(Result, 3); Result[0] := 1; Result[1] := 2; Result[2] := 3; end; var a: array of Integer; begin a := Make; WriteLn(a[2]); end."#
        ),
        &["3"]
    );
}

#[test]
fn dynamic_array_passed_to_open_array_param() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sum(const a: array of Integer): Integer; var i: Integer; begin Result := 0; for i := Low(a) to High(a) do Result := Result + a[i]; end; var data: array of Integer; begin data := [10, 20, 30]; WriteLn(Sum(data)); end."#
        ),
        &["60"]
    );
}

#[test]
fn dynamic_array_var_param_mutation_visible() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure DoubleAll(var a: array of Integer); var i: Integer; begin for i := Low(a) to High(a) do a[i] := a[i] * 2; end; var data: array of Integer; begin data := [3, 4]; DoubleAll(data); WriteLn(data[0]); WriteLn(data[1]); end."#
        ),
        &["6", "8"]
    );
}

#[test]
fn dynamic_array_setlength_zero_clears_length() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin a := [1, 2, 3]; SetLength(a, 0); WriteLn(Length(a)); WriteLn(High(a)); end."#
        ),
        &["0", "-1"]
    );
}

#[test]
fn dynamic_array_enum_element_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor = (Red, Green, Blue); var colors: array of TColor; begin colors := [Red, Blue]; WriteLn(Ord(colors[1])); end."#
        ),
        &["2"]
    );
}

#[test]
fn dynamic_array_nested_row_storage() {
    assert_eq!(
        run_pascal(
            r#"program T; var row1, row2: array of Integer; begin row1 := [1, 2]; row2 := [3, 4, 5]; WriteLn(Length(row1)); WriteLn(row2[2]); end."#
        ),
        &["2", "5"]
    );
}

#[test]
fn dynamic_array_append_via_setlength_and_assign() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin SetLength(a, 1); a[0] := 100; SetLength(a, 2); a[1] := 200; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["100", "200"]
    );
}
