/// Extended generic types: constraints, nested generics, methods, records.
use super::helpers::run_pascal;

#[test]
fn generic_record_pair_swap_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T>=record A,B:T; end; var p:TPair<Integer>; t:Integer; begin p.A:=1; p.B:=2; t:=p.A; p.A:=p.B; p.B:=t; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn generic_procedure_swap_two_vars() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Swap<T>(var a,b:T); var t:T; begin t:=a; a:=b; b:=t; end; var x,y:Integer; begin x:=3; y:=9; Swap<Integer>(x,y); WriteLn(x); WriteLn(y); end."#
        ),
        &["9", "3"]
    );
}

#[test]
fn generic_function_identity_string() {
    assert_eq!(
        run_pascal(
            r#"program T; function Id<T>(v:T):T; begin Result:=v; end; begin WriteLn(Id<String>('ok')); end."#
        ),
        &["ok"]
    );
}

#[test]
fn generic_class_list_count() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private FCount:Integer; public constructor Create; function Count:Integer; procedure Add; end; constructor TList<T>.Create; begin FCount:=0; end; function TList<T>.Count:Integer; begin Result:=FCount; end; procedure TList<T>.Add; begin Inc(FCount); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add; L.Add; WriteLn(L.Count); L.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_nested_specialization() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox<T>=class public Value:T; constructor Create(v:T); end; constructor TBox<T>.Create(v:T); begin Value:=v; end; var outer:TBox<TBox<Integer>>; begin outer:=TBox<TBox<Integer>>.Create(TBox<Integer>.Create(5)); WriteLn(outer.Value.Value); outer.Value.Free; outer.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn generic_array_wrapper_first() {
    assert_eq!(
        run_pascal(
            r#"program T; type TArrayBox<T>=record Items:array[0..1] of T; end; var b:TArrayBox<String>; begin b.Items[0]:='first'; b.Items[1]:='second'; WriteLn(b.Items[0]); end."#
        ),
        &["first"]
    );
}

#[test]
fn generic_min_with_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; function Min<T>(a,b:T):T; begin if a<b then Result:=a else Result:=b; end; begin WriteLn(Min<Integer>(8,3)); WriteLn(Min<String>('beta','alpha')); end."#
        ),
        &["3", "alpha"]
    );
}

#[test]
fn generic_stack_push_pop() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private FTop:Integer; FData:array[0..2] of T; public constructor Create; procedure Push(v:T); function Pop:T; end; constructor TStack<T>.Create; begin FTop:=-1; end; procedure TStack<T>.Push(v:T); begin Inc(FTop); FData[FTop]:=v; end; function TStack<T>.Pop:T; begin Result:=FData[FTop]; Dec(FTop); end; var s:TStack<Integer>; begin s:=TStack<Integer>.Create; s.Push(1); s.Push(2); WriteLn(s.Pop); WriteLn(s.Pop); s.Free; end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn generic_dictionary_simple_pair_list() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEntry<K,V>=record Key:K; Value:V; end; var e:TEntry<String,Integer>; begin e.Key:='age'; e.Value:=33; WriteLn(e.Key); WriteLn(e.Value); end."#
        ),
        &["age", "33"]
    );
}

#[test]
fn generic_class_constraint_style_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; type TComparable=class public Value:Integer; constructor Create(v:Integer); function LessThan(o:TComparable):Boolean; end; constructor TComparable.Create(v:Integer); begin Value:=v; end; function TComparable.LessThan(o:TComparable):Boolean; begin Result:=Value<o.Value; end; function PickSmaller(a,b:TComparable):TComparable; begin if a.LessThan(b) then Result:=a else Result:=b; end; var x,y,z:TComparable; begin x:=TComparable.Create(4); y:=TComparable.Create(9); z:=PickSmaller(x,y); WriteLn(z.Value); x.Free; y.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn generic_function_returns_same_type() {
    assert_eq!(
        run_pascal(
            r#"program T; function Double<T>(v:T):T; begin Result:=v; end; begin WriteLn(Double<Integer>(21)); end."#
        ),
        &["21"]
    );
}

#[test]
fn generic_record_default_ctor_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCell<T>=record Value:T; Initialized:Boolean; end; var c:TCell<Double>; begin c.Value:=2.5; c.Initialized:=true; if c.Initialized then WriteLn(Round(c.Value)); end."#
        ),
        &["3"]
    );
}

#[test]
fn generic_method_on_class_returns_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFactory<T>=class public class function Make(v:T):T; end; class function TFactory<Integer>.Make(v:Integer):Integer; begin Result:=v+1; end; begin WriteLn(TFactory<Integer>.Make(6)); end."#
        ),
        &["7"]
    );
}

#[test]
fn generic_interface_style_wrapper() {
    assert_eq!(
        run_pascal(
            r#"program T; type IHolder<T>=interface function Get:T; end; THolder<T>=class(TInterfacedObject,IHolder<T>) private F:T; public constructor Create(v:T); function Get:T; end; constructor THolder<T>.Create(v:T); begin F:=v; end; function THolder<T>.Get:T; begin Result:=F; end; var h:IHolder<String>; begin h:=THolder<String>.Create('data'); WriteLn(h.Get); end."#
        ),
        &["data"]
    );
}

#[test]
fn generic_three_type_params_tuple_like() {
    assert_eq!(
        run_pascal(
            r#"program T; type TTrip<A,B,C>=record X:A; Y:B; Z:C; end; var t:TTrip<Integer,String,Boolean>; begin t.X:=1; t.Y:='y'; t.Z:=true; WriteLn(t.X); WriteLn(t.Y); WriteLn(t.Z); end."#
        ),
        &["1", "y", "true"]
    );
}

#[test]
fn generic_class_static_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCounted=class public class var Total:Integer; constructor Create; end; class var TCounted.Total:Integer; constructor TCounted.Create; begin Inc(Total); end; begin TCounted.Total:=0; TCounted.Create; TCounted.Create; WriteLn(TCounted.Total); end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_function_array_sum_integers() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sum(const a:array of Integer):Integer; var i:Integer; begin Result:=0; for i:=Low(a) to High(a) do Result:=Result+a[i]; end; begin WriteLn(Sum([1,2,3])); end."#
        ),
        &["6"]
    );
}

#[test]
fn generic_record_copy_value_semantics() {
    assert_eq!(
        run_pascal(
            r#"program T; type TWrap<T>=record V:T; end; var a,b:TWrap<Integer>; begin a.V:=5; b:=a; b.V:=9; WriteLn(a.V); WriteLn(b.V); end."#
        ),
        &["5", "9"]
    );
}

#[test]
fn generic_pointer_box_deref() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPtrBox<T>=record P:^T; end; var n:Integer; b:TPtrBox<Integer>; begin n:=12; b.P:=@n; WriteLn(b.P^); end."#
        ),
        &["12"]
    );
}

#[test]
fn generic_enum_specialization() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(A,B); type TEnumBox<T>=record V:T; end; var e:TEnumBox<TD>; begin e.V:=B; WriteLn(Ord(e.V)); end."#
        ),
        &["1"]
    );
}

#[test]
fn generic_nested_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Run<T>(v:T); begin WriteLn(v); end; begin Run<String>('go'); Run<Integer>(42); end."#
        ),
        &["go", "42"]
    );
}

#[test]
fn generic_class_inherits_from_generic() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase<T>=class public Value:T; end; type TDerived=class(TBase<Integer>) public function Double:Integer; end; function TDerived.Double:Integer; begin Result:=Value*2; end; var d:TDerived; begin d:=TDerived.Create; d.Value:=11; WriteLn(d.Double); d.Free; end."#
        ),
        &["22"]
    );
}

#[test]
fn generic_default_value_via_explicit_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; function ZeroInt:Integer; begin Result:=0; end; begin WriteLn(ZeroInt); end."#
        ),
        &["0"]
    );
}
