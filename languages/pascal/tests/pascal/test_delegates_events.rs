/// Procedural types as events, multicast patterns, method pointers.
use super::helpers::run_pascal;

#[test]
fn event_handler_procedure_type_stored() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNotify=procedure(Sender: TObject); var h:TNotify; begin h:=procedure(Sender: TObject) begin WriteLn('fired'); end; h(nil); end."#
        ),
        &["fired"]
    );
}

#[test]
fn event_handler_with_integer_payload() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntEvent=procedure(v:Integer); var e:TIntEvent; begin e:=procedure(v:Integer) begin WriteLn(v); end; e(42); end."#
        ),
        &["42"]
    );
}

#[test]
fn event_handler_reassign_changes_behavior() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; var e:TEv; begin e:=procedure begin WriteLn('a'); end; e; e:=procedure begin WriteLn('b'); end; e; end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn event_multicast_two_handlers_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; procedure RunBoth(a,b:TEv); begin a; b; end; begin RunBoth(procedure begin WriteLn('1'); end, procedure begin WriteLn('2'); end); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn event_method_pointer_on_object() {
    assert_eq!(
        run_pascal(
            r#"program T; type TObj=class procedure Click; end; procedure TObj.Click; begin WriteLn('click'); end; var o:TObj; p:procedure of object; begin o:=TObj.Create; p:=o.Click; p; end."#
        ),
        &["click"]
    );
}

#[test]
fn event_method_pointer_with_state() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCounter=class FCount:Integer; procedure Inc; end; procedure TCounter.Inc; begin FCount:=FCount+1; end; var c:TCounter; p:procedure of object; begin c:=TCounter.Create; p:=c.Inc; p; p; WriteLn(c.FCount); end."#
        ),
        &["2"]
    );
}

#[test]
fn event_dispatch_table_two_methods() {
    assert_eq!(
        run_pascal(
            r#"program T; type THandler=class procedure A; procedure B; end; procedure THandler.A; begin WriteLn('A'); end; procedure THandler.B; begin WriteLn('B'); end; var h:THandler; p:procedure of object; begin h:=THandler.Create; p:=h.A; p; p:=h.B; p; end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn event_callback_passed_to_subscriber() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSub=class procedure Subscribe(cb:procedure); end; procedure TSub.Subscribe(cb:procedure); begin cb; end; var s:TSub; begin s:=TSub.Create; s.Subscribe(procedure begin WriteLn('sub'); end); end."#
        ),
        &["sub"]
    );
}

#[test]
fn event_filter_predicate_delegate() {
    assert_eq!(
        run_pascal(
            r#"program T; function Keep(n:Integer; pred:function(x:Integer):Boolean):Boolean; begin Result:=pred(n); end; begin WriteLn(Keep(4, function(x:Integer):Boolean begin Result:=x mod 2=0; end)); end."#
        ),
        &["true"]
    );
}

#[test]
fn event_map_delegate_over_array() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Each(fn:procedure(i:Integer)); var k:Integer; begin for k:=1 to 3 do fn(k); end; begin Each(procedure(i:Integer) begin WriteLn(i); end); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn event_handler_string_argument() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStrEv=procedure(const s:string); var e:TStrEv; begin e:=procedure(const s:string) begin WriteLn(s); end; e('evt'); end."#
        ),
        &["evt"]
    );
}

#[test]
fn event_chained_two_anonymous_calls() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Chain(a,b:procedure); begin a; b; end; begin Chain(procedure begin WriteLn('x'); end, procedure begin WriteLn('y'); end); end."#
        ),
        &["x", "y"]
    );
}

#[test]
fn event_method_pointer_virtual_dispatch() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class procedure Go; virtual; end; TChild=class(TBase) procedure Go; override; end; procedure TBase.Go; begin WriteLn('base'); end; procedure TChild.Go; begin WriteLn('child'); end; var b:TBase; p:procedure of object; begin b:=TChild.Create; p:=b.Go; p; end."#
        ),
        &["child"]
    );
}

#[test]
fn event_notify_list_simulated_with_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBus=class FHits:Integer; procedure Emit; end; procedure TBus.Emit; begin Inc(FHits); end; var b:TBus; p:procedure of object; begin b:=TBus.Create; p:=b.Emit; p; p; WriteLn(b.FHits); end."#
        ),
        &["2"]
    );
}

#[test]
fn event_compare_delegate_as_parameter() {
    assert_eq!(
        run_pascal(
            r#"program T; function Pick(a,b:Integer; cmp:function(x,y:Integer):Boolean):Integer; begin if cmp(a,b) then Result:=a else Result:=b; end; begin WriteLn(Pick(3,7, function(x,y:Integer):Boolean begin Result:=x>y; end)); end."#
        ),
        &["7"]
    );
}

#[test]
fn event_handler_nil_guard_skips() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; procedure SafeCall(e:TEv); begin if Assigned(e) then e; end; begin SafeCall(nil); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn event_handler_assigned_check_true() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; var e:TEv; begin e:=procedure begin WriteLn('yes'); end; if Assigned(e) then e; end."#
        ),
        &["yes"]
    );
}

#[test]
fn event_multicast_via_array_of_procedures() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; procedure FireAll(const hs:array of TEv); var i:Integer; begin for i:=Low(hs) to High(hs) do hs[i]; end; begin FireAll([procedure begin WriteLn('a'); end, procedure begin WriteLn('b'); end]); end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn event_on_change_simulated_property() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCell=class FV:Integer; FOnChange:procedure; procedure SetV(v:Integer); end; procedure TCell.SetV(v:Integer); begin FV:=v; if Assigned(FOnChange) then FOnChange; end; var c:TCell; begin c:=TCell.Create; c.FOnChange:=procedure begin WriteLn('changed'); end; c.SetV(1); end."#
        ),
        &["changed"]
    );
}

#[test]
fn event_timer_tick_simulation() {
    assert_eq!(
        run_pascal(
            r#"program T; type TTimer=class FTicks:Integer; FOnTick:procedure; procedure Run(n:Integer); var i:Integer; begin for i:=1 to n do if Assigned(FOnTick) then FOnTick; end; end; var t:TTimer; begin t:=TTimer.Create; t.FOnTick:=procedure begin Inc(t.FTicks); end; t.Run(3); WriteLn(t.FTicks); end."#
        ),
        &["3"]
    );
}

#[test]
fn event_method_pointer_returns_via_side_effect() {
    assert_eq!(
        run_pascal(
            r#"program T; type TAcc=class FSum:Integer; procedure Add(n:Integer); end; procedure TAcc.Add(n:Integer); begin FSum:=FSum+n; end; var a:TAcc; p:procedure of object; begin a:=TAcc.Create; p:=procedure begin a.Add(5); end; p; WriteLn(a.FSum); end."#
        ),
        &["5"]
    );
}

#[test]
fn event_double_dispatch_interface_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type IClick=interface procedure Click; end; TBtn=class(TInterfacedObject,IClick) procedure Click; end; procedure TBtn.Click; begin WriteLn('btn'); end; procedure Bind(c:IClick); begin c.Click; end; var b:IClick; begin b:=TBtn.Create; Bind(b); end."#
        ),
        &["btn"]
    );
}

#[test]
fn event_handler_capture_outer_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; var hits:Integer; p:procedure; begin hits:=0; p:=procedure begin Inc(hits); end; p; p; WriteLn(hits); end."#
        ),
        &["2"]
    );
}

#[test]
fn event_fold_with_delegate() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fold(start:Integer; fn:function(acc,x:Integer):Integer):Integer; begin Result:=fn(start,0); end; begin WriteLn(Fold(10, function(acc,x:Integer):Integer begin Result:=acc+5; end)); end."#
        ),
        &["15"]
    );
}

#[test]
fn event_procedure_of_object_in_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TObj=class procedure Ping; end; procedure TObj.Ping; begin WriteLn('ping'); end; type TSlot=record H:procedure of object; end; var o:TObj; s:TSlot; begin o:=TObj.Create; s.H:=o.Ping; s.H; end."#
        ),
        &["ping"]
    );
}

#[test]
fn event_multicast_three_listeners() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; procedure All(a,b,c:TEv); begin a; b; c; end; begin All(procedure begin WriteLn('1'); end, procedure begin WriteLn('2'); end, procedure begin WriteLn('3'); end); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn event_handler_bool_result_delegate() {
    assert_eq!(
        run_pascal(
            r#"program T; function Test(n:Integer; ok:function(x:Integer):Boolean):Boolean; begin Result:=ok(n); end; begin WriteLn(Test(10, function(x:Integer):Boolean begin Result:=x>5; end)); end."#
        ),
        &["true"]
    );
}

#[test]
fn event_unsubscribe_by_clearing_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSrc=class FOn:procedure; procedure Fire; end; procedure TSrc.Fire; begin if Assigned(FOn) then FOn; end; var s:TSrc; begin s:=TSrc.Create; s.FOn:=procedure begin WriteLn('once'); end; s.Fire; s.FOn:=nil; s.Fire; end."#
        ),
        &["once"]
    );
}

#[test]
fn event_method_pointer_inherited_method() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase=class procedure Run; virtual; end; TChild=class(TBase); procedure TBase.Run; begin WriteLn('run'); end; var c:TChild; p:procedure of object; begin c:=TChild.Create; p:=c.Run; p; end."#
        ),
        &["run"]
    );
}

#[test]
fn event_anonymous_with_local_string_capture() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure ShowMsg(const m:string); var p:procedure; begin p:=procedure begin WriteLn(m); end; p; end; begin ShowMsg('hello'); end."#
        ),
        &["hello"]
    );
}

#[test]
fn event_sort_comparator_delegate() {
    assert_eq!(
        run_pascal(
            r#"program T; function Earlier(a,b:Integer; less:function(x,y:Integer):Boolean):Boolean; begin Result:=less(a,b); end; begin WriteLn(Earlier(2,5, function(x,y:Integer):Boolean begin Result:=x<y; end)); end."#
        ),
        &["true"]
    );
}

#[test]
fn event_click_counter_on_object() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBtn=class FClicks:Integer; procedure Click; end; procedure TBtn.Click; begin Inc(FClicks); end; var b:TBtn; h:procedure of object; begin b:=TBtn.Create; h:=b.Click; h; h; WriteLn(b.FClicks); end."#
        ),
        &["2"]
    );
}

#[test]
fn event_raise_on_handler_invocation() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; var raised:Boolean; p:TEv; begin raised:=false; p:=procedure begin raised:=true; end; p; WriteLn(raised); end."#
        ),
        &["true"]
    );
}

#[test]
fn event_pass_method_pointer_to_helper() {
    assert_eq!(
        run_pascal(
            r#"program T; type TWorker=class procedure Do; end; procedure TWorker.Do; begin WriteLn('do'); end; procedure Invoke(p:procedure of object); begin p; end; var w:TWorker; begin w:=TWorker.Create; Invoke(w.Do); end."#
        ),
        &["do"]
    );
}

#[test]
fn event_two_objects_different_handlers() {
    assert_eq!(
        run_pascal(
            r#"program T; type TA=class procedure A; end; TB=class procedure B; end; procedure TA.A; begin WriteLn('A'); end; procedure TB.B; begin WriteLn('B'); end; var a:TA; b:TB; pa,pb:procedure of object; begin a:=TA.Create; b:=TB.Create; pa:=a.A; pb:=b.B; pa; pb; end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn event_delegate_stored_in_class_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type THost=class F:procedure; procedure Run; end; procedure THost.Run; begin if Assigned(F) then F; end; var h:THost; begin h:=THost.Create; h.F:=procedure begin WriteLn('go'); end; h.Run; end."#
        ),
        &["go"]
    );
}

#[test]
fn event_loop_break_on_delegate_result() {
    assert_eq!(
        run_pascal(
            r#"program T; function FindFirst(pred:function(i:Integer):Boolean):Integer; var i:Integer; begin Result:=-1; for i:=1 to 5 do if pred(i) then begin Result:=i; Exit; end; end; begin WriteLn(FindFirst(function(i:Integer):Boolean begin Result:=i=3; end)); end."#
        ),
        &["3"]
    );
}

#[test]
fn event_multicast_count_handlers() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEv=procedure; function CountCalls(const hs:array of TEv):Integer; var i,c:Integer; begin c:=0; for i:=Low(hs) to High(hs) do begin hs[i]; Inc(c); end; Result:=c; end; begin WriteLn(CountCalls([procedure begin end, procedure begin end])); end."#
        ),
        &["2"]
    );
}

#[test]
fn event_handler_with_two_integer_args() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPairEv=procedure(a,b:Integer); var e:TPairEv; begin e:=procedure(a,b:Integer) begin WriteLn(a+b); end; e(3,4); end."#
        ),
        &["7"]
    );
}

#[test]
fn event_method_pointer_free_after_call_safe() {
    assert_eq!(
        run_pascal(
            r#"program T; type TObj=class procedure Done; end; procedure TObj.Done; begin WriteLn('done'); end; var o:TObj; p:procedure of object; begin o:=TObj.Create; p:=o.Done; p; o.Free; WriteLn('ok'); end."#
        ),
        &["done", "ok"]
    );
}
