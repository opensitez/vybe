/// =, <>, <, >, <=, >= comparisons across types.
use super::helpers::run_pascal;

#[test]
fn int_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 5=5 then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn int_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 5<>3 then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn int_less() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 2<3 then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn int_greater() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 4>3 then WriteLn('gt'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn int_less_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 3<=3 then WriteLn('le'); end."#
        ),
        &["le"]
    );
}

#[test]
fn int_greater_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 7>=5 then WriteLn('ge'); end."#
        ),
        &["ge"]
    );
}

#[test]
fn real_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 2.5=2.5 then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn real_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 1.1<>2.2 then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn real_less() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 1.5<2.0 then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn real_greater() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 3.2>3.1 then WriteLn('gt'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn char_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Char; begin a:='x'; b:='x'; if a=b then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn char_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Char; begin a:='a'; b:='b'; if a<>b then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn char_less() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 'a'<'b' then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn char_greater() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 'z'>'a' then WriteLn('gt'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn string_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='hi'; if s='hi' then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn string_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='hi'; if s<>'bye' then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn string_less() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 'abc'<'abd' then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn string_greater() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 'zzz'>'aaa' then WriteLn('gt'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn bool_equal_true() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if true=true then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn bool_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if true<>false then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn enum_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(A,B); var x,y:T; begin x:=A; y:=A; if x=y then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn enum_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(A,B); var x,y:T; begin x:=A; y:=B; if x<>y then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn enum_less() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(A,B,C); var a,b:T; begin a:=A; b:=B; if a<b then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn enum_greater() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(A,B,C); var a,b:T; begin a:=C; b:=B; if a>b then WriteLn('gt'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn subrange_compare_le() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..10; var a,b:TR; begin a:=3; b:=5; if a<=b then WriteLn('le'); end."#
        ),
        &["le"]
    );
}

#[test]
fn subrange_compare_ge() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=1..10; var a,b:TR; begin a:=8; b:=5; if a>=b then WriteLn('ge'); end."#
        ),
        &["ge"]
    );
}

#[test]
fn negative_int_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if -5<-3 then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn zero_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 0=0 then WriteLn('eq'); end."#
        ),
        &["eq"]
    );
}

#[test]
fn mixed_int_compare_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=6; if (n>5) and (n<7) then WriteLn('mid'); end."#
        ),
        &["mid"]
    );
}

#[test]
fn compare_in_case() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=4; case n of 1..3:WriteLn('low'); 4..6:WriteLn('mid'); else WriteLn('hi'); end; end."#
        ),
        &["mid"]
    );
}

#[test]
fn string_empty_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:=''; if s='' then WriteLn('empty'); end."#
        ),
        &["empty"]
    );
}

#[test]
fn char_ord_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if Ord('B')>Ord('A') then WriteLn('gt'); end."#
        ),
        &["gt"]
    );
}

#[test]
fn real_le() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 2.0<=2.5 then WriteLn('le'); end."#
        ),
        &["le"]
    );
}

#[test]
fn real_ge() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 9.9>=9.8 then WriteLn('ge'); end."#
        ),
        &["ge"]
    );
}

#[test]
fn int_chain_ne() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if not (10=11) then WriteLn('ne'); end."#
        ),
        &["ne"]
    );
}

#[test]
fn compare_bool_in_if() {
    assert_eq!(
        run_pascal(
            r#"program T; var ok:Boolean; begin ok:=5>3; if ok then WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn string_ge_length() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if Length('abcd')>=4 then WriteLn('ge'); end."#
        ),
        &["ge"]
    );
}

#[test]
fn char_le_boundary() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if 'A'<='Z' then WriteLn('le'); end."#
        ),
        &["le"]
    );
}

#[test]
fn enum_le_ge() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(X,Y,Z); var v:T; begin v:=Y; if (v>=X) and (v<=Z) then WriteLn('range'); end."#
        ),
        &["range"]
    );
}

#[test]
fn pointer_equal_nil_style() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin p:=nil; if p=nil then WriteLn('nil'); end."#
        ),
        &["nil"]
    );
}

#[test]
fn set_compare_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of 1..3; begin s:=[]; if s=[] then WriteLn('empty'); end."#
        ),
        &["empty"]
    );
}

#[test]
fn multi_type_compare_vars() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; r:Real; begin i:=5; r:=5.0; if i=Trunc(r) then WriteLn('match'); end."#
        ),
        &["match"]
    );
}

