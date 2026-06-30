/// Scoped enumerations (`enum TypeName: underlying`) in modern Delphi.
use super::helpers::run_pascal;

#[test]
fn scoped_enum_byte_underlying_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStatus = (Off, On); var s: TStatus; begin s := On; WriteLn(Ord(s)); end."#
        ),
        &["1"]
    );
}

#[test]
fn scoped_enum_explicit_values() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCode = (Ok = 0, Warn = 10, Err = 20); var c: TCode; begin c := Warn; WriteLn(Ord(c)); end."#
        ),
        &["10"]
    );
}

#[test]
fn scoped_enum_case_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLevel = (Low, Mid, High); var l: TLevel; begin l := Mid; case l of Low: WriteLn('L'); Mid: WriteLn('M'); High: WriteLn('H'); end; end."#
        ),
        &["M"]
    );
}

#[test]
fn scoped_enum_succ_pred_walk() {
    assert_eq!(
        run_pascal(
            r#"program T; type T = (A, B, C); var v: T; begin v := A; v := Succ(v); WriteLn(Ord(v)); v := Pred(Succ(v)); WriteLn(Ord(v)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn scoped_enum_for_in_array_of_enums() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay = (Mon, Tue, Wed); var d: TDay; n: Integer; begin n := 0; for d in [Mon, Wed] do n := n + 1; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn scoped_enum_set_membership() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFlag = (A, B, C); var chosen: set of TFlag; begin chosen := [B, C]; if B in chosen then WriteLn('has B'); end."#
        ),
        &["has B"]
    );
}

#[test]
fn scoped_enum_record_field_storage() {
    assert_eq!(
        run_pascal(
            r#"program T; type TKind = (Cat, Dog); type TPet = record Kind: TKind; Name: String; end; var p: TPet; begin p.Kind := Dog; p.Name := 'Rex'; WriteLn(p.Name); WriteLn(Ord(p.Kind)); end."#
        ),
        &["Rex", "1"]
    );
}

#[test]
fn scoped_enum_function_returns_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMode = (Read, Write); function DefaultMode: TMode; begin Result := Read; end; begin WriteLn(Ord(DefaultMode)); end."#
        ),
        &["0"]
    );
}
