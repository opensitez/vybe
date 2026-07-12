/// If/else chains, case else, nested case — distinct from test_case_statement_ranges.rs.
use super::helpers::run_pascal;

#[test]
fn if_elseif_else_middle_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=2; if x=1 then WriteLn('one') else if x=2 then WriteLn('two') else WriteLn('other'); end."#
        ),
        &["two"]
    );
}

#[test]
fn if_elseif_else_final_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=9; if x=1 then WriteLn('one') else if x=2 then WriteLn('two') else WriteLn('other'); end."#
        ),
        &["other"]
    );
}

#[test]
fn if_elseif_three_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='B'; if s='A' then WriteLn('alpha') else if s='B' then WriteLn('beta') else if s='C' then WriteLn('gamma') else WriteLn('?'); end."#
        ),
        &["beta"]
    );
}

#[test]
fn if_without_else_false_branch() {
    assert_eq!(
        run_pascal(r#"program T; begin if false then WriteLn('no'); WriteLn('yes'); end."#),
        &["yes"]
    );
}

#[test]
fn if_not_condition() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:Boolean; begin f:=false; if not f then WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn if_and_short_circuit_second_false() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if (1<2) and (3>4) then WriteLn('yes') else WriteLn('no'); end."#
        ),
        &["no"]
    );
}

#[test]
fn if_or_short_circuit_second_true() {
    assert_eq!(
        run_pascal(
            r#"program T; begin if (1>2) or (3<4) then WriteLn('yes') else WriteLn('no'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn if_nested_three_deep() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c:Integer; begin a:=1; b:=1; c:=0; if a=1 then if b=1 then if c=0 then WriteLn('deep'); end."#
        ),
        &["deep"]
    );
}

#[test]
fn if_comparison_chain_greater_equal() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=10; if n>=10 then WriteLn('ge'); end."#),
        &["ge"]
    );
}

#[test]
fn if_string_equality() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='pascal'; if s='pascal' then WriteLn('match'); end."#
        ),
        &["match"]
    );
}

#[test]
fn if_char_compare() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin c:='Z'; if c='Z' then WriteLn('z'); end."#),
        &["z"]
    );
}

#[test]
fn case_simple_integer_label() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; case n of 1:WriteLn('a'); 2:WriteLn('b'); 3:WriteLn('c'); end; end."#
        ),
        &["b"]
    );
}

#[test]
fn case_else_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=99; case n of 1:WriteLn('one'); else WriteLn('other'); end; end."#
        ),
        &["other"]
    );
}

#[test]
fn case_no_match_without_else_skips() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=5; case n of 1,2,3:WriteLn('hit'); end; WriteLn('after'); end."#
        ),
        &["after"]
    );
}

#[test]
fn case_char_labels() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='b'; case c of 'a':WriteLn('A'); 'b':WriteLn('B'); else WriteLn('?'); end; end."#
        ),
        &["B"]
    );
}

#[test]
fn case_boolean_as_integer_zero_one() {
    assert_eq!(
        run_pascal(
            r#"program T; var b:Boolean; begin b:=true; case Ord(b) of 0:WriteLn('f'); 1:WriteLn('t'); end; end."#
        ),
        &["t"]
    );
}

#[test]
fn case_enum_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TS=(Low,High); var s:TS; begin s:=High; case s of Low:WriteLn('l'); High:WriteLn('h'); end; end."#
        ),
        &["h"]
    );
}

#[test]
fn case_nested_inner_match() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=1; b:=2; case a of 1:case b of 2:WriteLn('ok'); else WriteLn('no'); end; else WriteLn('skip'); end; end."#
        ),
        &["ok"]
    );
}

#[test]
fn case_nested_inner_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=1; b:=9; case a of 1:case b of 2:WriteLn('two'); else WriteLn('inner'); end; else WriteLn('outer'); end; end."#
        ),
        &["inner"]
    );
}

#[test]
fn case_statement_sets_variable() {
    assert_eq!(
        run_pascal(
            r#"program T; var n,code:Integer; begin n:=3; code:=0; case n of 1:code:=10; 2:code:=20; 3:code:=30; end; WriteLn(code); end."#
        ),
        &["30"]
    );
}

#[test]
fn case_with_begin_end_block() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=1; case n of 1:begin WriteLn('a'); WriteLn('b'); end; end; end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn if_in_function_early_result() {
    assert_eq!(
        run_pascal(
            r#"program T; function Classify(n:Integer):string; begin if n<0 then Result:='neg' else if n=0 then Result:='zero' else Result:='pos'; end; begin WriteLn(Classify(-1)); WriteLn(Classify(0)); WriteLn(Classify(2)); end."#
        ),
        &["neg", "zero", "pos"]
    );
}

#[test]
fn case_in_function_with_else() {
    assert_eq!(
        run_pascal(
            r#"program T; function Label(n:Integer):string; begin case n of 0:Result:='zero'; 1:Result:='one'; else Result:='many'; end; end; begin WriteLn(Label(1)); WriteLn(Label(8)); end."#
        ),
        &["one", "many"]
    );
}

#[test]
fn if_empty_then_part_skipped() {
    assert_eq!(
        run_pascal(r#"program T; begin if false then ; WriteLn('next'); end."#),
        &["next"]
    );
}

#[test]
fn case_multiple_labels_same_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:Integer; begin d:=2; case d of 1,2,3:WriteLn('small'); 4,5:WriteLn('big'); end; end."#
        ),
        &["small"]
    );
}

#[test]
fn if_modulo_even_odd() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=7; if n mod 2=0 then WriteLn('even') else WriteLn('odd'); end."#
        ),
        &["odd"]
    );
}

#[test]
fn case_negative_integer_label() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=-1; case n of -1:WriteLn('neg'); 0:WriteLn('zero'); else WriteLn('other'); end; end."#
        ),
        &["neg"]
    );
}

#[test]
fn nested_if_else_flips_outer() {
    assert_eq!(
        run_pascal(
            r#"program T; var x,y:Integer; begin x:=0; y:=1; if x=1 then WriteLn('a') else if y=1 then WriteLn('b') else WriteLn('c'); end."#
        ),
        &["b"]
    );
}

#[test]
fn case_on_string_first_char() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; c:Char; begin s:='test'; c:=s[1]; case c of 't':WriteLn('t'); else WriteLn('?'); end; end."#
        ),
        &["t"]
    );
}

#[test]
fn if_in_loop_filters_output() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin for i:=1 to 4 do if i mod 2=1 then WriteLn(i); end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn case_in_loop_dispatch() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin for i:=1 to 3 do case i of 1:WriteLn('a'); 2:WriteLn('b'); 3:WriteLn('c'); end; end."#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn if_elseif_on_char_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='M'; if c<'N' then WriteLn('before') else if c='N' then WriteLn('n') else WriteLn('after'); end."#
        ),
        &["before"]
    );
}

#[test]
fn case_else_in_nested_outer_miss() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=9; b:=1; case a of 1:case b of 1:WriteLn('in'); end; else WriteLn('out'); end; end."#
        ),
        &["out"]
    );
}

#[test]
fn if_boolean_literal_and_expression() {
    assert_eq!(
        run_pascal(r#"program T; begin if true and (2>1) then WriteLn('ok'); end."#),
        &["ok"]
    );
}

#[test]
fn case_zero_label() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; case n of 0:WriteLn('zero'); else WriteLn('nz'); end; end."#
        ),
        &["zero"]
    );
}

#[test]
fn if_elseif_first_true_short() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=1; if n=1 then WriteLn('first') else if n=2 then WriteLn('second') else WriteLn('third'); end."#
        ),
        &["first"]
    );
}

#[test]
fn case_two_sequential_statements() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; case n of 1:WriteLn('one'); end; case n of 2:WriteLn('two'); end; end."#
        ),
        &["two"]
    );
}

#[test]
fn if_xor_condition() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Boolean; begin a:=true; b:=false; if a xor b then WriteLn('xor'); end."#
        ),
        &["xor"]
    );
}

#[test]
fn case_enum_with_else_via_integer_tag() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=(Red,Green,Blue); var n:Integer; begin n:=9; case TC(n) of Red:WriteLn('r'); Green:WriteLn('g'); Blue:WriteLn('b'); else WriteLn('?'); end; end."#
        ),
        &["?"]
    );
}

#[test]
fn if_nested_case_mixed() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; if n>0 then case n of 1:WriteLn('one'); 2:WriteLn('two'); end else WriteLn('neg'); end."#
        ),
        &["two"]
    );
}

#[test]
fn case_assign_string_result() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; s:string; begin n:=5; case n of 5:s:='five'; else s:='?'; end; WriteLn(s); end."#
        ),
        &["five"]
    );
}
