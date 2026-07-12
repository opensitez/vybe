/// Repeat/until edge cases, nested if, case without ranges.
use super::helpers::run_pascal;

#[test]
fn repeat_until_counter_hits_exact_bound() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; repeat Inc(n); until n=7; WriteLn(n); end."#
        ),
        &["7"]
    );
}

#[test]
fn repeat_until_condition_true_after_first_body() {
    assert_eq!(
        run_pascal(
            r#"program T; var done:Boolean; begin done:=false; repeat done:=true; until done; WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn repeat_until_nested_inner_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,s:Integer; begin s:=0; i:=0; repeat Inc(i); j:=0; repeat Inc(j); s:=s+1; until j=2; until i=2; WriteLn(s); end."#
        ),
        &["4"]
    );
}

#[test]
fn repeat_until_outer_wraps_if() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; repeat Inc(n); if n=3 then WriteLn('hit'); until n>=3; end."#
        ),
        &["hit"]
    );
}

#[test]
fn repeat_until_not_equal_exit() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=5; repeat Dec(x); until not (x>0); WriteLn(x); end."#
        ),
        &["0"]
    );
}

#[test]
fn repeat_until_xor_two_flags() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=0; b:=1; repeat Inc(a); until (a>2) xor (b=0); WriteLn(a); end."#
        ),
        &["3"]
    );
}

#[test]
fn repeat_until_modulo_cycle() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; repeat Inc(n); until (n mod 5)=0; WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn repeat_until_string_length_target() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:=''; repeat s:=s+'*'; until Length(s)=6; WriteLn(Length(s)); end."#
        ),
        &["6"]
    );
}

#[test]
fn repeat_until_body_updates_condition_var() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=1; b:=10; repeat b:=b-a; until b<=0; WriteLn(b); end."#
        ),
        &["0"]
    );
}

#[test]
fn repeat_until_many_small_steps() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; repeat Inc(i); until i=12; WriteLn(i); end."#
        ),
        &["12"]
    );
}

#[test]
fn repeat_until_with_begin_block_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var n,s:Integer; begin n:=0; s:=0; repeat begin Inc(n); s:=s+n; end; until n=4; WriteLn(s); end."#
        ),
        &["10"]
    );
}

#[test]
fn repeat_until_decrement_hits_target() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=5; repeat Dec(n); until n=2; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_if_three_levels_all_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=9; if x>0 then if x>5 then if x>8 then WriteLn('deep') else WriteLn('mid') else WriteLn('shallow'); end."#
        ),
        &["deep"]
    );
}

#[test]
fn nested_if_middle_branch_false() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=3; if x>0 then if x>5 then WriteLn('a') else WriteLn('b') else WriteLn('c'); end."#
        ),
        &["b"]
    );
}

#[test]
fn nested_if_outer_false_skips_inner() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=-1; if x>0 then if x>5 then WriteLn('a') else WriteLn('b') else WriteLn('c'); end."#
        ),
        &["c"]
    );
}

#[test]
fn nested_if_else_binds_to_innermost() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; if n>0 then if n=1 then WriteLn('one') else WriteLn('not-one'); end."#
        ),
        &["not-one"]
    );
}

#[test]
fn nested_if_with_and_condition() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; begin a:=3; b:=4; if a>0 then if (a<5) and (b>3) then WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn nested_if_with_or_shortcut() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=0; if x=0 then if (x=1) or (x=0) then WriteLn('zero'); end."#
        ),
        &["zero"]
    );
}

#[test]
fn nested_if_not_negation() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:Boolean; begin f:=false; if not f then if not false then WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn nested_if_four_levels_deepest() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; begin v:=4; if v>0 then if v>1 then if v>2 then if v>3 then WriteLn('4') else WriteLn('3'); end."#
        ),
        &["4"]
    );
}

#[test]
fn nested_if_blocks_multiple_statements() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=2; if x>1 then begin if x<5 then begin WriteLn('a'); WriteLn('b'); end; end; end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn nested_if_comparison_chain_greater() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=50; if n>10 then if n>20 then if n>30 then WriteLn('high'); end."#
        ),
        &["high"]
    );
}

#[test]
fn nested_if_inside_repeat_body() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; repeat Inc(i); if i=2 then WriteLn('two'); until i=3; end."#
        ),
        &["two"]
    );
}

#[test]
fn case_integer_single_value_hit() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=4; case n of 2: WriteLn('two'); 4: WriteLn('four'); else WriteLn('other'); end; end."#
        ),
        &["four"]
    );
}

#[test]
fn case_integer_single_value_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=99; case n of 1: WriteLn('one'); 2: WriteLn('two'); else WriteLn('else'); end; end."#
        ),
        &["else"]
    );
}

#[test]
fn case_integer_zero_label() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=0; case n of 0: WriteLn('zero'); 1: WriteLn('one'); end; end."#
        ),
        &["zero"]
    );
}

#[test]
fn case_integer_negative_labels() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=-3; case n of -1: WriteLn('a'); -3: WriteLn('b'); else WriteLn('c'); end; end."#
        ),
        &["b"]
    );
}

#[test]
fn case_char_single_letter_match() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='k'; case c of 'a': WriteLn('a'); 'k': WriteLn('k'); else WriteLn('z'); end; end."#
        ),
        &["k"]
    );
}

#[test]
fn case_char_no_match_falls_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='q'; case c of 'a': WriteLn('a'); 'b': WriteLn('b'); else WriteLn('other'); end; end."#
        ),
        &["other"]
    );
}

#[test]
fn case_char_digit_labels() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='5'; case c of '0','1','2','3','4': WriteLn('low'); '5','6','7','8','9': WriteLn('five-plus'); end; end."#
        ),
        &["five-plus"]
    );
}

#[test]
fn case_multiple_labels_same_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=3; case n of 1,3,5: WriteLn('odd'); 2,4,6: WriteLn('even'); end; end."#
        ),
        &["odd"]
    );
}

#[test]
fn case_separate_labels_different_actions() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; case n of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."#
        ),
        &["two"]
    );
}

#[test]
fn case_with_begin_block_body() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=1; case n of 1: begin WriteLn('a'); WriteLn('b'); end; end; end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn case_inside_if_true_guard() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; if n>0 then case n of 1: WriteLn('one'); 2: WriteLn('two'); end; end."#
        ),
        &["two"]
    );
}

#[test]
fn case_inside_repeat_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; repeat Inc(i); case i of 1: WriteLn('a'); 2: WriteLn('b'); end; until i=2; end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn case_else_only_path_taken() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=100; case n of 1: WriteLn('one'); else WriteLn('fallback'); end; end."#
        ),
        &["fallback"]
    );
}

#[test]
fn case_two_arm_only_first() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=1; case n of 1: WriteLn('first'); 2: WriteLn('second'); end; end."#
        ),
        &["first"]
    );
}

#[test]
fn case_two_arm_only_second() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; case n of 1: WriteLn('first'); 2: WriteLn('second'); end; end."#
        ),
        &["second"]
    );
}

#[test]
fn repeat_flag_simulated_break() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,st:Integer; begin i:=0; st:=0; repeat Inc(i); if i=3 then st:=1; until st=1; WriteLn(i); end."#
        ),
        &["3"]
    );
}

#[test]
fn nested_if_equality_on_chars() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='x'; if c='x' then if c<>'y' then WriteLn('match'); end."#
        ),
        &["match"]
    );
}

#[test]
fn repeat_until_sum_exceeds_threshold() {
    assert_eq!(
        run_pascal(
            r#"program T; var k,s:Integer; begin k:=0; s:=0; repeat Inc(k); s:=s+k; until s>20; WriteLn(k); end."#
        ),
        &["6"]
    );
}

#[test]
fn case_enum_ordinal_discrete() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor=(Red,Green,Blue); var c:TColor; begin c:=Green; case c of Red: WriteLn('r'); Green: WriteLn('g'); Blue: WriteLn('b'); end; end."#
        ),
        &["g"]
    );
}
