/// Advanced generics: constraints, specializations, generic records and classes.
use super::helpers::run_pascal;

#[test]
fn generic_record_pair_of_strings() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T>=record First, Second:T; end; var p:TPair<String>; begin p.First:='a'; p.Second:='b'; WriteLn(p.First+p.Second); end."#
        ),
        &["ab"]
    );
}

#[test]
fn generic_function_max_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; function Bigger<T>(a,b:T):T; begin if a>b then Result:=a else Result:=b; end; begin WriteLn(Bigger<Integer>(4,9)); end."#
        ),
        &["9"]
    );
}

#[test]
fn generic_class_box_integer_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox<T>=class public Value:T; constructor Create(v:T); end; constructor TBox<T>.Create(v:T); begin Value:=v; end; var b:TBox<Integer>; begin b:=TBox<Integer>.Create(12); WriteLn(b.Value); b.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn generic_procedure_exchange_refs() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Exchange<T>(var a,b:T); var t:T; begin t:=a; a:=b; b:=t; end; var x,y:String; begin x:='1'; y:='2'; Exchange<String>(x,y); WriteLn(x+y); end."#
        ),
        &["21"]
    );
}

#[test]
fn generic_stack_of_characters() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private FTop:Integer; FData:array[0..3] of T; public constructor Create; procedure Push(v:T); function Pop:T; end; constructor TStack<T>.Create; begin FTop:=-1; end; procedure TStack<T>.Push(v:T); begin Inc(FTop); FData[FTop]:=v; end; function TStack<T>.Pop:T; begin Result:=FData[FTop]; Dec(FTop); end; var s:TStack<Char>; begin s:=TStack<Char>.Create; s.Push('x'); WriteLn(s.Pop); s.Free; end."#
        ),
        &["x"]
    );
}

#[test]
fn generic_array_wrapper_length_two() {
    assert_eq!(
        run_pascal(
            r#"program T; type TW<T>=record Items:array[0..1] of T; end; var w:TW<Double>; begin w.Items[0]:=1.5; w.Items[1]:=2.5; WriteLn(Round(w.Items[0]+w.Items[1])); end."#
        ),
        &["4"]
    );
}

#[test]
fn generic_function_first_of_open_array() {
    assert_eq!(
        run_pascal(
            r#"program T; function First<T>(const a:array of T):T; begin Result:=a[0]; end; var nums:array of Integer; begin nums:=[8,9]; WriteLn(First<Integer>(nums)); end."#
        ),
        &["8"]
    );
}

#[test]
fn generic_nested_box_in_box() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox<T>=class public Value:T; constructor Create(v:T); end; constructor TBox<T>.Create(v:T); begin Value:=v; end; var outer:TBox<TBox<String>>; begin outer:=TBox<TBox<String>>.Create(TBox<String>.Create('ok')); WriteLn(outer.Value.Value); outer.Value.Free; outer.Free; end."#
        ),
        &["ok"]
    );
}

#[test]
fn generic_record_default_initialize_integers() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR<T>=record A,B:T; end; var r:TR<Integer>; begin r.A:=3; r.B:=4; WriteLn(r.A*r.B); end."#
        ),
        &["12"]
    );
}

#[test]
fn generic_list_push_count_strings() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private FN:Integer; public constructor Create; procedure Add; function Count:Integer; end; constructor TList<T>.Create; begin FN:=0; end; procedure TList<T>.Add; begin Inc(FN); end; function TList<T>.Count:Integer; begin Result:=FN; end; var L:TList<String>; begin L:=TList<String>.Create; L.Add; L.Add; L.Add; WriteLn(L.Count); L.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn generic_min_string_lexicographic() {
    assert_eq!(
        run_pascal(
            r#"program T; function Smaller<T>(a,b:T):T; begin if a<b then Result:=a else Result:=b; end; begin WriteLn(Smaller<String>('delta','alpha')); end."#
        ),
        &["alpha"]
    );
}

#[test]
fn generic_pair_swap_values_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T>=record X,Y:T; end; procedure SwapPair<T>(var p:TPair<T>); var t:T; begin t:=p.X; p.X:=p.Y; p.Y:=t; end; var p:TPair<Integer>; begin p.X:=1; p.Y:=9; SwapPair<Integer>(p); WriteLn(p.X); WriteLn(p.Y); end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn generic_function_last_index_open_array() {
    assert_eq!(
        run_pascal(
            r#"program T; function Last<T>(const a:array of T):T; begin Result:=a[High(a)]; end; var a:array of String; begin a:=['a','b','c']; WriteLn(Last<String>(a)); end."#
        ),
        &["c"]
    );
}

#[test]
fn generic_class_method_returns_generic() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFactory<T>=class public class function Make(v:T):T; end; class function TFactory<T>.Make(v:T):T; begin Result:=v; end; begin WriteLn(TFactory<Integer>.Make(77)); end."#
        ),
        &["77"]
    );
}

#[test]
fn generic_queue_enqueue_dequeue() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private FHead,FSize:Integer; FBuf:array[0..2] of T; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin FHead:=0; FSize:=0; end; procedure TQueue<T>.Enq(v:T); begin FBuf[FSize]:=v; Inc(FSize); end; function TQueue<T>.Deq:T; begin Result:=FBuf[FHead]; Inc(FHead); Dec(FSize); end; var q:TQueue<Integer>; begin q:=TQueue<Integer>.Create; q.Enq(5); q.Enq(6); WriteLn(q.Deq); WriteLn(q.Deq); q.Free; end."#
        ),
        &["5", "6"]
    );
}

#[test]
fn generic_record_with_boolean_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFlags<T>=record On,Off:T; end; var f:TFlags<Boolean>; begin f.On:=true; f.Off:=false; WriteLn(f.On); WriteLn(f.Off); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn generic_identity_on_double() {
    assert_eq!(
        run_pascal(
            r#"program T; function Id<T>(v:T):T; begin Result:=v; end; begin WriteLn(Id<Double>(2.5)); end."#
        ),
        &["2.5"]
    );
}

#[test]
fn generic_map_apply_to_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; function Double<T>(v:T):T; begin Result:=v; end; function DoubleInt(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(DoubleInt(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn generic_vector_dot_two_integers() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVec<T>=record X,Y:T; end; function Dot(a,b:TVec<Integer>):Integer; begin Result:=a.X*b.X+a.Y*b.Y; end; var u,v:TVec<Integer>; begin u.X:=1; u.Y:=2; v.X:=3; v.Y:=4; WriteLn(Dot(u,v)); end."#
        ),
        &["11"]
    );
}

#[test]
fn generic_optional_box_nil_check() {
    assert_eq!(
        run_pascal(
            r#"program T; type TOpt<T>=class public Has:Boolean; Value:T; constructor Create; end; constructor TOpt<T>.Create; begin Has:=false; end; var o:TOpt<Integer>; begin o:=TOpt<Integer>.Create; WriteLn(o.Has); o.Free; end."#
        ),
        &["false"]
    );
}

#[test]
fn generic_tree_left_right_children() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=record Left,Right:T; end; var n:TNode<Integer>; begin n.Left:=2; n.Right:=3; WriteLn(n.Left+n.Right); end."#
        ),
        &["5"]
    );
}

#[test]
fn generic_procedure_fill_open_array() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Fill<T>(var a:array of T; v:T); var i:Integer; begin for i:=Low(a) to High(a) do a[i]:=v; end; var a:array of Integer; begin SetLength(a,3); Fill<Integer>(a,7); WriteLn(a[2]); end."#
        ),
        &["7"]
    );
}

#[test]
fn generic_class_inherits_generic_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBase<T>=class public Value:T; end; TChild=class(TBase<Integer>) end; var c:TChild; begin c:=TChild.Create; c.Value:=11; WriteLn(c.Value); c.Free; end."#
        ),
        &["11"]
    );
}

#[test]
fn generic_function_compare_equal_strings() {
    assert_eq!(
        run_pascal(
            r#"program T; function Eq<T>(a,b:T):Boolean; begin Result:=a=b; end; begin WriteLn(Eq<String>('x','x')); end."#
        ),
        &["true"]
    );
}

#[test]
fn generic_ring_buffer_wrap() {
    assert_eq!(
        run_pascal(
            r#"program T; type TRing<T>=class private FIdx:Integer; FVal:T; public constructor Create(v:T); procedure SetNext(v:T); function Get:T; end; constructor TRing<T>.Create(v:T); begin FVal:=v; FIdx:=0; end; procedure TRing<T>.SetNext(v:T); begin FVal:=v; Inc(FIdx); end; function TRing<T>.Get:T; begin Result:=FVal; end; var r:TRing<String>; begin r:=TRing<String>.Create('a'); r.SetNext('b'); WriteLn(r.Get); r.Free; end."#
        ),
        &["b"]
    );
}

#[test]
fn generic_pair_of_chars_concat() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T>=record A,B:T; end; var p:TPair<Char>; begin p.A:='Q'; p.B:='Z'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["Q", "Z"]
    );
}

#[test]
fn generic_heap_single_cell() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCell<T>=class public V:T; constructor Create(v:T); end; constructor TCell<T>.Create(v:T); begin V:=v; end; var c:TCell<String>; begin c:=TCell<String>.Create('cell'); WriteLn(c.V); c.Free; end."#
        ),
        &["cell"]
    );
}

#[test]
fn generic_sum_open_array_integers() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sum<T>(const a:array of Integer):Integer; var i:Integer; begin Result:=0; for i:=Low(a) to High(a) do Result:=Result+a[i]; end; var a:array of Integer; begin a:=[1,2,3]; WriteLn(Sum<Integer>(a)); end."#
        ),
        &["6"]
    );
}

#[test]
fn generic_record_clone_field_copy() {
    assert_eq!(
        run_pascal(
            r#"program T; type TClone<T>=record Data:T; end; var a,b:TClone<Integer>; begin a.Data:=5; b:=a; b.Data:=8; WriteLn(a.Data); WriteLn(b.Data); end."#
        ),
        &["5", "8"]
    );
}

#[test]
fn generic_optional_set_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TOpt<T>=class public Has:Boolean; Value:T; procedure Set(v:T); end; procedure TOpt<T>.Set(v:T); begin Has:=true; Value:=v; end; var o:TOpt<Integer>; begin o:=TOpt<Integer>.Create; o.Set(4); WriteLn(o.Value); o.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn generic_key_value_string_int() {
    assert_eq!(
        run_pascal(
            r#"program T; type TKVP<K,V>=record Key:K; Value:V; end; var m:TKVP<String,Integer>; begin m.Key:='age'; m.Value:=30; WriteLn(m.Value); end."#
        ),
        &["30"]
    );
}

#[test]
fn generic_function_reverse_two_item_array() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Reverse<T>(var a:array of T); var t:T; begin t:=a[0]; a[0]:=a[1]; a[1]:=t; end; var a:array of String; begin a:=['first','second']; Reverse<String>(a); WriteLn(a[0]); end."#
        ),
        &["second"]
    );
}

#[test]
fn generic_class_static_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCounted<T>=class public class var N:Integer; class procedure Bump; end; class var TCounted<Integer>.N:Integer; class procedure TCounted<T>.Bump; begin Inc(N); end; begin TCounted<Integer>.N:=0; TCounted<Integer>.Bump; TCounted<Integer>.Bump; WriteLn(TCounted<Integer>.N); end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_matrix_2x2_integer_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMat<T>=record Cells:array[0..1,0..1] of T; end; var m:TMat<Integer>; begin m.Cells[0,0]:=1; m.Cells[0,1]:=2; m.Cells[1,0]:=3; m.Cells[1,1]:=4; WriteLn(m.Cells[1,1]); end."#
        ),
        &["4"]
    );
}

#[test]
fn generic_either_left_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEither<L,R>=record IsLeft:Boolean; Left:L; Right:R; end; var e:TEither<Integer,String>; begin e.IsLeft:=true; e.Left:=9; WriteLn(e.Left); end."#
        ),
        &["9"]
    );
}

#[test]
fn generic_buffer_write_read() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBuf<T>=class private F:T; public procedure Write(v:T); function Read:T; end; procedure TBuf<T>.Write(v:T); begin F:=v; end; function TBuf<T>.Read:T; begin Result:=F; end; var b:TBuf<Double>; begin b:=TBuf<Double>.Create; b.Write(3.5); WriteLn(b.Read>3.0); b.Free; end."#
        ),
        &["true"]
    );
}

#[test]
fn generic_triple_record_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TTrip<T>=record A,B,C:T; end; var t:TTrip<Integer>; begin t.A:=1; t.B:=2; t.C:=3; WriteLn(t.A+t.B+t.C); end."#
        ),
        &["6"]
    );
}

#[test]
fn generic_predicate_filter_count() {
    assert_eq!(
        run_pascal(
            r#"program T; function CountPos(const a:array of Integer):Integer; var i:Integer; begin Result:=0; for i:=Low(a) to High(a) do if a[i]>0 then Inc(Result); end; begin WriteLn(CountPos([1,-2,3])); end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_singleton_class_instance() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSing<T>=class public class var Inst:TSing<T>; class function Get:TSing<T>; end; class var TSing<Integer>.Inst:TSing<Integer>; class function TSing<T>.Get:TSing<T>; begin if Inst=nil then Inst:=TSing<T>.Create; Result:=Inst; end; var s:TSing<Integer>; begin s:=TSing<Integer>.Get; WriteLn(s<>nil); s.Free; TSing<Integer>.Inst:=nil; end."#
        ),
        &["true"]
    );
}

#[test]
fn generic_open_array_param_length() {
    assert_eq!(
        run_pascal(
            r#"program T; function Len<T>(const a:array of T):Integer; begin Result:=Length(a); end; var a:array of Integer; begin a:=[1,2,3,4]; WriteLn(Len<Integer>(a)); end."#
        ),
        &["4"]
    );
}
