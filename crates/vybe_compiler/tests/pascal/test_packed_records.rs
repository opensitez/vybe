/// Packed records and storage layout sensitive field access.
use super::helpers::run_pascal;

#[test]
fn packed_record_two_bytes_size() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPacked = packed record A: Byte; B: Byte; end; var p: TPacked; begin p.A := 1; p.B := 2; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn packed_record_mixed_byte_char_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMix = packed record Code: Byte; Ch: Char; end; var m: TMix; begin m.Code := 65; m.Ch := 'Z'; WriteLn(m.Ch); end."#
        ),
        &["Z"]
    );
}

#[test]
fn packed_record_assign_copy() {
    assert_eq!(
        run_pascal(
            r#"program T; type TP = packed record Lo, Hi: Byte; end; var a, b: TP; begin a.Lo := 3; a.Hi := 9; b := a; WriteLn(b.Hi); end."#
        ),
        &["9"]
    );
}

#[test]
fn packed_record_nested_in_regular_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner = packed record B: Byte; end; type TOuter = record Inner: TInner; Tag: Integer; end; var o: TOuter; begin o.Inner.B := 7; o.Tag := 1; WriteLn(o.Inner.B); end."#
        ),
        &["7"]
    );
}

#[test]
fn packed_record_used_in_case_on_byte_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type THdr = packed record Ver: Byte; end; var h: THdr; begin h.Ver := 2; case h.Ver of 1: WriteLn('v1'); 2: WriteLn('v2'); else WriteLn('other'); end; end."#
        ),
        &["v2"]
    );
}

#[test]
fn packed_record_boolean_byte_pair() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFlags = packed record On: Boolean; Id: Byte; end; var f: TFlags; begin f.On := true; f.Id := 5; if f.On then WriteLn(f.Id); end."#
        ),
        &["5"]
    );
}
