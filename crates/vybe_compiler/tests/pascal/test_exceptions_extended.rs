/// Exception handling: try/except/finally, raise, nested, and typed exceptions.
use super::helpers::run_pascal;

#[test]
fn try_except_catches_division_by_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=1; try x:=x div 0; except WriteLn('caught'); end; end."#
        ),
        &["caught"]
    );
}

#[test]
fn try_except_else_branch_not_taken() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('ok'); except WriteLn('bad'); end; end."#
        ),
        &["ok"]
    );
}

#[test]
fn try_finally_always_runs() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try WriteLn('try'); finally WriteLn('finally'); end; end."#
        ),
        &["try", "finally"]
    );
}

#[test]
fn try_finally_runs_after_except() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('e'); except WriteLn('ex'); end; finally WriteLn('fin'); end; end."#
        ),
        &["ex", "fin"]
    );
}

#[test]
fn raise_explicit_exception_message() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create('boom'); except on E:Exception do WriteLn(E.Message); end; end."#
        ),
        &["boom"]
    );
}

#[test]
fn except_on_specific_type_only() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise ERangeError.Create('range'); except on E:ERangeError do WriteLn('range'); on E:Exception do WriteLn('other'); end; end."#
        ),
        &["range"]
    );
}

#[test]
fn nested_try_blocks_inner_catch() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('inner'); except WriteLn('inner'); end; except WriteLn('outer'); end; end."#
        ),
        &["inner"]
    );
}

#[test]
fn try_except_sets_recovery_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; var ok:Boolean; begin ok:=false; try raise Exception.Create('x'); except ok:=true; end; WriteLn(ok); end."#
        ),
        &["true"]
    );
}

#[test]
fn finally_preserves_return_value_style() {
    assert_eq!(
        run_pascal(
            r#"program T; function F:Integer; begin Result:=0; try Result:=5; finally Result:=Result+1; end; end; begin WriteLn(F); end."#
        ),
        &["6"]
    );
}

#[test]
fn re_raise_after_log() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; try try raise Exception.Create('e'); except Inc(n); raise; end; except Inc(n); end; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn except_outer_catches_inner_reraise() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try raise Exception.Create('x'); except raise Exception.Create('y'); end; except on E:Exception do WriteLn(E.Message); end; end."#
        ),
        &["y"]
    );
}

#[test]
fn try_finally_with_break_inside() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin for i:=1 to 3 do try if i=2 then break; finally WriteLn('f'); end; end."#
        ),
        &["f", "f"]
    );
}

#[test]
fn exception_in_constructor_cleanup() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFail=class constructor Create; end; constructor TFail.Create; begin inherited; raise Exception.Create('ctor'); end; begin try TFail.Create; except WriteLn('ctor-fail'); end; end."#
        ),
        &["ctor-fail"]
    );
}

#[test]
fn assert_raises_in_debug_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try Assert(false); except WriteLn('assert'); end; end."#
        ),
        &["assert"]
    );
}

#[test]
fn array_index_out_of_range_caught() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[1..2] of Integer; i:Integer; begin a[1]:=1; a[2]:=2; try i:=a[9]; except WriteLn('bounds'); end; end."#
        ),
        &["bounds"]
    );
}

#[test]
fn try_except_continue_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,n:Integer; begin n:=0; for i:=1 to 3 do begin try if i=2 then raise Exception.Create('skip'); n:=n+1; except end; end; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn exception_class_inheritance_match() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise EDivByZero.Create('div'); except on E:EDivByZero do WriteLn('div0'); end; end."#
        ),
        &["div0"]
    );
}

#[test]
fn finally_counter_increments() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Integer; begin c:=0; try try raise Exception.Create('e'); except end; finally Inc(c); end; WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn nested_finally_order() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try try WriteLn('a'); finally WriteLn('b'); end; finally WriteLn('c'); end; end."#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn raise_empty_exception_still_caught() {
    assert_eq!(
        run_pascal(
            r#"program T; begin try raise Exception.Create(''); except on E:Exception do if E.Message='' then WriteLn('empty'); end; end."#
        ),
        &["empty"]
    );
}
