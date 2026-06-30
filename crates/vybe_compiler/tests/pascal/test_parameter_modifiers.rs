/// Parameter modifiers: var, const, out — distinct calling conventions in Delphi.
use super::helpers::run_pascal;

#[test]
fn var_parameter_mutates_caller_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure IncTwice(var n: Integer); begin n := n + 2; end; var x: Integer; begin x := 5; IncTwice(x); WriteLn(x); end."#
        ),
        &["7"]
    );
}

#[test]
fn var_parameter_swaps_two_variables() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Swap(var a, b: Integer); var t: Integer; begin t := a; a := b; b := t; end; var x, y: Integer; begin x := 1; y := 9; Swap(x, y); WriteLn(x); WriteLn(y); end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn const_parameter_cannot_be_assigned_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; function Twice(const n: Integer): Integer; begin Result := n * 2; end; begin WriteLn(Twice(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn const_string_parameter_passed_read_only() {
    assert_eq!(
        run_pascal(
            r#"program T; function Len(const s: String): Integer; begin Result := Length(s); end; begin WriteLn(Len('abcd')); end."#
        ),
        &["4"]
    );
}

#[test]
fn out_parameter_assigns_before_return() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Split(out a, b: Integer); begin a := 3; b := 4; end; var x, y: Integer; begin Split(x, y); WriteLn(x); WriteLn(y); end."#
        ),
        &["3", "4"]
    );
}

#[test]
fn out_parameter_uninitialized_on_entry() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure MakePair(out lo, hi: Integer); begin lo := 10; hi := 20; end; var a, b: Integer; begin MakePair(a, b); WriteLn(a + b); end."#
        ),
        &["30"]
    );
}

#[test]
fn var_open_array_parameter_mutates_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure NegateAll(var a: array of Integer); var i: Integer; begin for i := Low(a) to High(a) do a[i] := -a[i]; end; var data: array of Integer; begin data := [1, -2, 3]; NegateAll(data); WriteLn(data[0]); WriteLn(data[1]); end."#
        ),
        &["-1", "2"]
    );
}

#[test]
fn const_open_array_parameter_read_only_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; function Total(const a: array of Integer): Integer; var i: Integer; begin Result := 0; for i := Low(a) to High(a) do Result := Result + a[i]; end; begin WriteLn(Total([2, 4, 6])); end."#
        ),
        &["12"]
    );
}

#[test]
fn var_record_parameter_mutates_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = record N: Integer; end; procedure Bump(var r: TR); begin r.N := r.N + 1; end; var rec: TR; begin rec.N := 8; Bump(rec); WriteLn(rec.N); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_record_parameter_reads_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = record Name: String; end; function LabelOf(const r: TR): String; begin Result := '[' + r.Name + ']'; end; var rec: TR; begin rec.Name := 'x'; WriteLn(LabelOf(rec)); end."#
        ),
        &["[x]"]
    );
}

#[test]
fn multiple_var_parameters_updated_in_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Fill(var a, b, c: Integer); begin a := 1; b := 2; c := 3; end; var x, y, z: Integer; begin Fill(x, y, z); WriteLn(x); WriteLn(y); WriteLn(z); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn out_string_parameter_returns_concatenated_value() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Build(out s: String); begin s := 'built'; end; var msg: String; begin Build(msg); WriteLn(msg); end."#
        ),
        &["built"]
    );
}
