/// Repeat/until variants, for-step patterns, loop idioms beyond test_control_flow.rs.
use super::helpers::run_pascal;

#[test]
fn while_manual_step_by_two() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin i:=0; while i<=6 do begin WriteLn(i); i:=i+2; end; end."#),
        &["0", "2", "4", "6"]
    );
}

#[test]
fn while_manual_step_by_three_down() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin i:=9; while i>=0 do begin WriteLn(i); i:=i-3; end; end."#),
        &["9", "6", "3", "0"]
    );
}

#[test]
fn for_to_with_variable_end_bound() {
    assert_eq!(
        run_pascal(r#"program T; var i,n:Integer; begin n:=4; for i:=1 to n do WriteLn(i); end."#),
        &["1", "2", "3", "4"]
    );
}

#[test]
fn for_downto_with_variable_start() {
    assert_eq!(
        run_pascal(r#"program T; var i,n:Integer; begin n:=3; for i:=n downto 1 do WriteLn(i); end."#),
        &["3", "2", "1"]
    );
}

#[test]
fn repeat_until_sum_reaches_target() {
    assert_eq!(
        run_pascal(r#"program T; var n,s:Integer; begin n:=0; s:=0; repeat Inc(n); s:=s+n; until s>=10; WriteLn(s); end."#),
        &["10"]
    );
}

#[test]
fn repeat_until_string_length_met() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; begin s:=''; repeat s:=s+'x'; until Length(s)=4; WriteLn(Length(s)); end."#),
        &["4"]
    );
}

#[test]
fn repeat_until_false_first_iteration_runs() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=0; repeat WriteLn('body'); Inc(n); until n>0; end."#),
        &["body"]
    );
}

#[test]
fn repeat_until_with_or_condition() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Integer; begin a:=0; b:=10; repeat Inc(a); until (a>=3) or (b<0); WriteLn(a); end."#),
        &["3"]
    );
}

#[test]
fn repeat_until_with_and_condition() {
    assert_eq!(
        run_pascal(r#"program T; var x,y:Integer; begin x:=0; y:=0; repeat Inc(x); Inc(y); until (x>=2) and (y>=2); WriteLn(x); end."#),
        &["2"]
    );
}

#[test]
fn nested_for_row_column_matrix() {
    assert_eq!(
        run_pascal(r#"program T; var r,c:Integer; begin for r:=1 to 2 do for c:=1 to 2 do WriteLn(IntToStr(r)+','+IntToStr(c)); end."#),
        &["1,1", "1,2", "2,1", "2,2"]
    );
}

#[test]
fn nested_repeat_inner_break_outer() {
    assert_eq!(
        run_pascal(r#"program T; var i,j:Integer; begin i:=0; repeat Inc(i); j:=0; repeat Inc(j); if j=2 then Break; until false; until i=2; WriteLn(i); WriteLn(j); end."#),
        &["2", "2"]
    );
}

#[test]
fn for_empty_range_skips_body() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=5 to 3 do WriteLn(i); WriteLn('end'); end."#),
        &["end"]
    );
}

#[test]
fn for_downto_empty_range_skips() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=1 downto 5 do WriteLn(i); WriteLn('skip'); end."#),
        &["skip"]
    );
}

#[test]
fn while_continue_skip_even() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin i:=0; while i<5 do begin Inc(i); if i mod 2=0 then Continue; WriteLn(i); end; end."#),
        &["1", "3", "5"]
    );
}

#[test]
fn for_continue_skip_multiples_of_three() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=1 to 6 do begin if i mod 3=0 then Continue; WriteLn(i); end; end."#),
        &["1", "2", "4", "5"]
    );
}

#[test]
fn repeat_continue_in_middle() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=0; repeat Inc(n); if n=2 then Continue; WriteLn(n); until n>=4; end."#),
        &["1", "3", "4"]
    );
}

#[test]
fn for_break_on_match() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=1 to 10 do begin if i=4 then Break; WriteLn(i); end; end."#),
        &["1", "2", "3"]
    );
}

#[test]
fn while_break_on_condition() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=0; while true do begin Inc(n); if n=3 then Break; WriteLn(n); end; end."#),
        &["1", "2"]
    );
}

#[test]
fn repeat_nested_for_accumulator() {
    assert_eq!(
        run_pascal(r#"program T; var outer,inner,sum:Integer; begin sum:=0; outer:=0; repeat Inc(outer); inner:=0; for inner:=1 to outer do sum:=sum+1; until outer=3; WriteLn(sum); end."#),
        &["6"]
    );
}

#[test]
fn for_char_range_loop() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin for c:='A' to 'C' do WriteLn(c); end."#),
        &["A", "B", "C"]
    );
}

#[test]
fn for_char_downto_range() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin for c:='C' downto 'A' do WriteLn(c); end."#),
        &["C", "B", "A"]
    );
}

#[test]
fn while_decrement_to_zero() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=3; while n>0 do begin WriteLn(n); Dec(n); end; end."#),
        &["3", "2", "1"]
    );
}

#[test]
fn repeat_until_modulo_cycle() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=0; repeat Inc(n); until n mod 7=0; WriteLn(n); end."#),
        &["7"]
    );
}

#[test]
fn for_nested_break_inner_only() {
    assert_eq!(
        run_pascal(r#"program T; var i,j:Integer; begin for i:=1 to 3 do begin for j:=1 to 3 do begin if j=2 then Break; WriteLn(IntToStr(i)+'-'+IntToStr(j)); end; end; end."#),
        &["1-1", "2-1", "3-1"]
    );
}

#[test]
fn repeat_with_exit_procedure() {
    assert_eq!(
        run_pascal(r#"program T; procedure Scan; var n:Integer; begin n:=0; repeat Inc(n); if n=2 then Exit; WriteLn('loop'); until false; end; begin Scan; WriteLn('out'); end."#),
        &["loop", "out"]
    );
}

#[test]
fn for_loop_variable_preserved_after() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=1 to 3 do ; WriteLn(i); end."#),
        &["4"]
    );
}

#[test]
fn for_downto_loop_variable_after_exit() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=5 downto 1 do ; WriteLn(i); end."#),
        &["0"]
    );
}

#[test]
fn while_repeat_equivalent_count() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Integer; begin a:=0; while a<3 do begin Inc(a); end; b:=0; repeat Inc(b); until b>=3; WriteLn(a); WriteLn(b); end."#),
        &["3", "3"]
    );
}

#[test]
fn repeat_until_not_condition() {
    assert_eq!(
        run_pascal(r#"program T; var done:Boolean; begin done:=false; repeat WriteLn('tick'); done:=true; until not done; end."#),
        &["tick"]
    );
}

#[test]
fn for_step_simulated_with_while_positive() {
    assert_eq!(
        run_pascal(r#"program T; var i,start,stop,step:Integer; begin start:=2; stop:=10; step:=2; i:=start; while i<=stop do begin WriteLn(i); i:=i+step; end; end."#),
        &["2", "4", "6", "8", "10"]
    );
}

#[test]
fn for_step_simulated_with_while_negative() {
    assert_eq!(
        run_pascal(r#"program T; var i,start,stop,step:Integer; begin start:=10; stop:=2; step:=-2; i:=start; while i>=stop do begin WriteLn(i); i:=i+step; end; end."#),
        &["10", "8", "6", "4", "2"]
    );
}

#[test]
fn nested_while_triangular_number() {
    assert_eq!(
        run_pascal(r#"program T; var i,j,sum:Integer; begin sum:=0; i:=1; while i<=4 do begin j:=1; while j<=i do begin sum:=sum+1; Inc(j); end; Inc(i); end; WriteLn(sum); end."#),
        &["10"]
    );
}

#[test]
fn repeat_factorial_accumulation() {
    assert_eq!(
        run_pascal(r#"program T; var i,f:Integer; begin i:=1; f:=1; repeat f:=f*i; Inc(i); until i>5; WriteLn(f); end."#),
        &["120"]
    );
}

#[test]
fn for_enum_members_iteration() {
    assert_eq!(
        run_pascal(r#"program T; type TColor=(Red,Green,Blue); var c:TColor; begin for c:=Red to Blue do WriteLn(Ord(c)); end."#),
        &["0", "1", "2"]
    );
}

#[test]
fn repeat_until_char_sequence_match() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; begin s:=''; repeat s:=s+'a'; until s='aaa'; WriteLn(s); end."#),
        &["aaa"]
    );
}

#[test]
fn while_flag_toggle_twice() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; on:Boolean; begin n:=0; on:=true; while on do begin Inc(n); if n=2 then on:=false; WriteLn(n); end; end."#),
        &["1", "2"]
    );
}

#[test]
fn for_inner_break_does_not_stop_outer() {
    assert_eq!(
        run_pascal(r#"program T; var i,j,hits:Integer; begin hits:=0; for i:=1 to 2 do begin for j:=1 to 3 do begin Inc(hits); if j=1 then Break; end; end; WriteLn(hits); end."#),
        &["2"]
    );
}

#[test]
fn repeat_double_continue_pattern() {
    assert_eq!(
        run_pascal(r#"program T; var n,outc:Integer; begin n:=0; outc:=0; repeat Inc(n); if n mod 2=0 then Continue; if n>5 then Break; Inc(outc); until false; WriteLn(outc); end."#),
        &["3"]
    );
}

#[test]
fn for_to_high_bound_expression() {
    assert_eq!(
        run_pascal(r#"program T; var i,n:Integer; begin n:=2; for i:=1 to n+n do WriteLn(i); end."#),
        &["1", "2", "3", "4"]
    );
}

#[test]
fn repeat_until_zero_div_guard_simulated() {
    assert_eq!(
        run_pascal(r#"program T; var n,safe:Integer; begin n:=5; safe:=0; repeat if n=0 then Break; safe:=100 div n; Dec(n); until n=0; WriteLn(safe); end."#),
        &["25"]
    );
}
