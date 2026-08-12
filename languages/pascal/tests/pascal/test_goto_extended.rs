/// Goto/label patterns beyond test_goto_label.rs.
use super::helpers::run_pascal;

#[test]
fn goto_backward_label_loop_count() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var n:Integer; begin n:=0; top: Inc(n); if n<3 then goto top; WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn goto_forward_over_multiple_statements() {
    assert_eq!(
        run_pascal(
            r#"program T; label finish; begin WriteLn('start'); goto finish; WriteLn('skip1'); WriteLn('skip2'); finish: WriteLn('done'); end."#
        ),
        &["start", "done"]
    );
}

#[test]
fn goto_conditional_branch_merge() {
    assert_eq!(
        run_pascal(
            r#"program T; label merge; var x:Integer; begin x:=1; if x=1 then goto merge; WriteLn('no'); merge: WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn goto_nested_if_escape() {
    assert_eq!(
        run_pascal(
            r#"program T; label out_label; var a,b:Integer; begin a:=0; b:=1; if a=1 then begin if b=1 then WriteLn('in'); end else goto out_label; WriteLn('fall'); out_label: WriteLn('out'); end."#
        ),
        &["out"]
    );
}

#[test]
fn goto_while_replacement_count_down() {
    assert_eq!(
        run_pascal(
            r#"program T; label again; var n:Integer; begin n:=3; again: WriteLn(n); Dec(n); if n>0 then goto again; end."#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn goto_error_handler_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T; label err, ok; var code:Integer; begin code:=-1; if code<0 then goto err; WriteLn('ok'); goto ok; err: WriteLn('err'); ok: WriteLn('end'); end."#
        ),
        &["err", "end"]
    );
}

#[test]
fn goto_skip_else_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; label past; var n:Integer; begin n:=5; if n>0 then goto past; WriteLn('neg'); past: WriteLn('past'); end."#
        ),
        &["past"]
    );
}

#[test]
fn goto_two_labels_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; label L1, L2; begin WriteLn('a'); goto L1; L2: WriteLn('c'); goto done; L1: WriteLn('b'); goto L2; done: end."#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn goto_break_outer_loop_simulation() {
    assert_eq!(
        run_pascal(
            r#"program T; label break_outer; var i,j:Integer; begin for i:=1 to 3 do begin for j:=1 to 3 do begin if (i=2) and (j=2) then goto break_outer; end; end; break_outer: WriteLn('broken'); end."#
        ),
        &["broken"]
    );
}

#[test]
fn goto_continue_loop_simulation() {
    assert_eq!(
        run_pascal(
            r#"program T; label loop_top, loop_cont; var i:Integer; begin i:=0; loop_top: Inc(i); if i=2 then goto loop_cont; WriteLn(i); loop_cont: if i<3 then goto loop_top; end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn goto_case_like_dispatch() {
    assert_eq!(
        run_pascal(
            r#"program T; label L1,L2,L3,done; var n:Integer; begin n:=2; if n=1 then goto L1 else if n=2 then goto L2 else goto L3; L1: WriteLn('one'); goto done; L2: WriteLn('two'); goto done; L3: WriteLn('three'); done: end."#
        ),
        &["two"]
    );
}

#[test]
fn goto_procedure_exit_early() {
    assert_eq!(
        run_pascal(
            r#"program T; label ret; procedure P(n:Integer); begin if n=0 then goto ret; WriteLn(n); ret: end; begin P(0); P(2); end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_repeat_until_alternative() {
    assert_eq!(
        run_pascal(
            r#"program T; label check; var n:Integer; begin n:=0; check: WriteLn(n); Inc(n); if n<=2 then goto check; end."#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn goto_forward_declaration_before_use() {
    assert_eq!(
        run_pascal(r#"program T; label target; begin goto target; target: WriteLn('hit'); end."#),
        &["hit"]
    );
}

#[test]
fn goto_skip_initialization_block() {
    assert_eq!(
        run_pascal(
            r#"program T; label run; var ready:Boolean; begin ready:=false; if not ready then goto run; WriteLn('init'); run: WriteLn('run'); end."#
        ),
        &["run"]
    );
}

#[test]
fn goto_chain_three_hops() {
    assert_eq!(
        run_pascal(
            r#"program T; label A,B,C; begin goto A; C: WriteLn('c'); goto done; B: goto C; A: goto B; done: end."#
        ),
        &["c"]
    );
}

#[test]
fn goto_with_inc_accumulator() {
    assert_eq!(
        run_pascal(
            r#"program T; label again; var sum,i:Integer; begin sum:=0; i:=0; again: Inc(i); sum:=sum+i; if i<4 then goto again; WriteLn(sum); end."#
        ),
        &["10"]
    );
}

#[test]
fn goto_boolean_flag_exit() {
    assert_eq!(
        run_pascal(
            r#"program T; label exit_loop; var i:Integer; done:Boolean; begin i:=0; done:=false; while not done do begin Inc(i); if i=2 then done:=true; if done then goto exit_loop; WriteLn('tick'); end; exit_loop: WriteLn('out'); end."#
        ),
        &["tick", "out"]
    );
}

#[test]
fn goto_skip_second_write() {
    assert_eq!(
        run_pascal(
            r#"program T; label endblock; begin WriteLn('first'); goto endblock; WriteLn('second'); endblock: end."#
        ),
        &["first"]
    );
}

#[test]
fn goto_nested_procedure_label() {
    assert_eq!(
        run_pascal(
            r#"program T; label done; procedure Inner; label leave; begin goto leave; WriteLn('skip'); leave: WriteLn('inner'); end; begin Inner; goto done; WriteLn('skip2'); done: WriteLn('outer'); end."#
        ),
        &["inner", "outer"]
    );
}

#[test]
fn goto_for_loop_escape() {
    assert_eq!(
        run_pascal(
            r#"program T; label bail; var i:Integer; begin for i:=1 to 10 do begin if i=4 then goto bail; WriteLn(i); end; bail: WriteLn('stop'); end."#
        ),
        &["1", "2", "3", "stop"]
    );
}

#[test]
fn goto_if_false_fallthrough() {
    assert_eq!(
        run_pascal(
            r#"program T; label L; begin if false then goto L; WriteLn('fall'); L: WriteLn('L'); end."#
        ),
        &["fall", "L"]
    );
}

#[test]
fn goto_multiple_entry_single_exit() {
    assert_eq!(
        run_pascal(
            r#"program T; label exit_label; var r:Integer; begin r:=1; if r=1 then goto exit_label; if r=2 then goto exit_label; WriteLn('nope'); exit_label: WriteLn('exit'); end."#
        ),
        &["exit"]
    );
}

#[test]
fn goto_string_compare_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; label match, nomatch; var s:string; begin s:='ok'; if s='ok' then goto match else goto nomatch; match: WriteLn('match'); goto done; nomatch: WriteLn('no'); done: end."#
        ),
        &["match"]
    );
}

#[test]
fn goto_decrement_until_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var n:Integer; begin n:=2; top: WriteLn(n); Dec(n); if n>=0 then goto top; end."#
        ),
        &["2", "1", "0"]
    );
}

#[test]
fn goto_repeat_body_once_then_jump() {
    assert_eq!(
        run_pascal(
            r#"program T; label once; begin WriteLn('body'); goto once; once: WriteLn('once'); end."#
        ),
        &["body", "once"]
    );
}

#[test]
fn goto_cross_if_blocks() {
    assert_eq!(
        run_pascal(
            r#"program T; label target; var a:Integer; begin a:=0; if a=0 then begin goto target; end; WriteLn('skip'); target: WriteLn('tgt'); end."#
        ),
        &["tgt"]
    );
}

#[test]
fn goto_simulate_case_else() {
    assert_eq!(
        run_pascal(
            r#"program T; label L1,L2,Lelse,done; var n:Integer; begin n:=9; if n=1 then goto L1 else if n=2 then goto L2 else goto Lelse; L1: WriteLn('1'); goto done; L2: WriteLn('2'); goto done; Lelse: WriteLn('else'); done: end."#
        ),
        &["else"]
    );
}

#[test]
fn goto_finally_block_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T; label cleanup; var ok:Boolean; begin ok:=false; if not ok then goto cleanup; WriteLn('work'); cleanup: WriteLn('cleanup'); end."#
        ),
        &["cleanup"]
    );
}

#[test]
fn goto_char_scan_stop_on_match() {
    assert_eq!(
        run_pascal(
            r#"program T; label found; var i:Integer; s:string; c:Char; begin s:='abc'; for i:=1 to Length(s) do begin c:=s[i]; if c='b' then goto found; end; found: WriteLn(c); end."#
        ),
        &["b"]
    );
}

#[test]
fn goto_reset_and_retry_once() {
    assert_eq!(
        run_pascal(
            r#"program T; label retry, giveup; var tries:Integer; begin tries:=0; retry: Inc(tries); if tries=1 then goto retry; if tries>2 then goto giveup; WriteLn(tries); giveup: WriteLn('done'); end."#
        ),
        &["2", "done"]
    );
}

#[test]
fn goto_skip_nested_begin_block() {
    assert_eq!(
        run_pascal(
            r#"program T; label out_label; begin begin goto out_label; WriteLn('in'); end; out_label: WriteLn('out'); end."#
        ),
        &["out"]
    );
}

#[test]
fn goto_modulo_filter_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; label top; var n:Integer; begin n:=0; top: Inc(n); if (n mod 2=0) or (n<5) then goto top; WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn goto_two_separate_forward_labels() {
    assert_eq!(
        run_pascal(
            r#"program T; label A,B; begin goto A; B: WriteLn('B'); goto endprog; A: WriteLn('A'); goto B; endprog: end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn goto_procedure_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Step1; label step2; begin goto step2; WriteLn('skip'); step2: WriteLn('step2'); end; begin Step1; end."#
        ),
        &["step2"]
    );
}

#[test]
fn goto_compare_integers_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; label gt,le; var a,b:Integer; begin a:=5; b:=3; if a>b then goto gt else goto le; gt: WriteLn('gt'); goto endprog; le: WriteLn('le'); endprog: end."#
        ),
        &["gt"]
    );
}

#[test]
fn goto_exit_before_second_label() {
    assert_eq!(
        run_pascal(
            r#"program T; label L1,L2; begin goto L2; L1: WriteLn('1'); L2: WriteLn('2'); end."#
        ),
        &["2"]
    );
}

#[test]
fn goto_while_true_break_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T; label break_loop; var n:Integer; begin n:=0; while true do begin Inc(n); if n=2 then goto break_loop; WriteLn(n); end; break_loop: WriteLn('brk'); end."#
        ),
        &["1", "brk"]
    );
}

#[test]
fn goto_accumulate_until_threshold() {
    assert_eq!(
        run_pascal(
            r#"program T; label add_more; var sum,v:Integer; begin sum:=0; v:=1; add_more: sum:=sum+v; Inc(v); if sum<10 then goto add_more; WriteLn(sum); end."#
        ),
        &["10"]
    );
}

#[test]
fn goto_skip_after_success_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; label fail, ok; var success:Boolean; begin success:=true; if success then goto ok; fail: WriteLn('fail'); goto endprog; ok: WriteLn('ok'); endprog: end."#
        ),
        &["ok"]
    );
}

#[test]
fn goto_three_level_nested_escape() {
    assert_eq!(
        run_pascal(
            r#"program T; label escape; var depth:Integer; begin depth:=3; if depth=3 then if depth>0 then goto escape; WriteLn('no'); escape: WriteLn('out'); end."#
        ),
        &["out"]
    );
}
