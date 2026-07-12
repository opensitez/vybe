/// Dictionary and collection-style patterns via generics and records.
use super::helpers::run_pascal;

#[test]
fn manual_map_record_lookup() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEntry=record Key:string; Value:Integer; end; var items:array[0..1] of TEntry; function Get(const k:string):Integer; var i:Integer; begin Result:=-1; for i:=0 to 1 do if items[i].Key=k then Result:=items[i].Value; end; begin items[0].Key:='a'; items[0].Value:=1; items[1].Key:='b'; items[1].Value:=2; WriteLn(Get('b')); end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_list_push_pop() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; FTop:Integer; public constructor Create; procedure Push(v:T); function Pop:T; end; constructor TStack<T>.Create; begin FTop:=-1; end; procedure TStack<T>.Push(v:T); begin Inc(FTop); SetLength(F,FTop+1); F[FTop]:=v; end; function TStack<T>.Pop:T; begin Result:=F[FTop]; Dec(FTop); end; var s:TStack<Integer>; begin s:=TStack<Integer>.Create; s.Push(1); s.Push(2); WriteLn(s.Pop); s.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_queue_fifo() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead,FTail:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin FHead:=0; FTail:=-1; end; procedure TQueue<T>.Enq(v:T); begin Inc(FTail); SetLength(F,FTail+1); F[FTail]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var q:TQueue<string>; begin q:=TQueue<string>.Create; q.Enq('first'); q.Enq('second'); WriteLn(q.Deq); q.Free; end."#
        ),
        &["first"]
    );
}

#[test]
fn set_as_unique_collection() {
    assert_eq!(
        run_pascal(
            r#"program T; var seen:set of Integer; n:Integer; c:Integer; begin seen:=[1,2,3]; c:=0; for n in seen do Inc(c); WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn dynamic_array_as_list_append() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin SetLength(a,1); a[0]:=1; SetLength(a,2); a[1]:=2; WriteLn(a[0]+a[1]); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_dictionary_two_keys() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDict=record K1,K2:Integer; end; var d:TDict; begin d.K1:=10; d.K2:=20; WriteLn(d.K1+d.K2); end."#
        ),
        &["30"]
    );
}

#[test]
fn array_of_pairs_iterate() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record K:string; V:Integer; end; var p:array[0..1] of TPair; i:Integer; s:Integer; begin p[0].K:='x'; p[0].V:=1; p[1].K:='y'; p[1].V:=2; s:=0; for i:=0 to 1 do s:=s+p[i].V; WriteLn(s); end."#
        ),
        &["3"]
    );
}

#[test]
fn generic_key_value_box() {
    assert_eq!(
        run_pascal(
            r#"program T; type TKV<K,V>=record Key:K; Value:V; end; var e:TKV<string,Integer>; begin e.Key:='id'; e.Value:=42; WriteLn(e.Value); end."#
        ),
        &["42"]
    );
}

#[test]
fn list_contains_linear_search() {
    assert_eq!(
        run_pascal(
            r#"program T; function Contains(const a:array of Integer; v:Integer):Boolean; var i:Integer; begin Result:=false; for i:=Low(a) to High(a) do if a[i]=v then Result:=true; end; begin WriteLn(Contains([1,3,5],3)); WriteLn(Contains([1,3,5],9)); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn map_update_existing_key() {
    assert_eq!(
        run_pascal(
            r#"program T; type TEntry=record Key:string; Value:Integer; end; var items:array[0..0] of TEntry; procedure Put(k:string; v:Integer); begin items[0].Key:=k; items[0].Value:=v; end; begin Put('x',1); Put('x',9); WriteLn(items[0].Value); end."#
        ),
        &["9"]
    );
}

#[test]
fn collection_count_with_for_in() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; x,c:Integer; begin a:=[4,5,6]; c:=0; for x in a do Inc(c); WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn multiset_count_duplicates_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; i,c:Integer; begin a:=[1,1,2]; c:=0; for i:=0 to High(a) do if a[i]=1 then Inc(c); WriteLn(c); end."#
        ),
        &["2"]
    );
}

#[test]
fn string_list_join() {
    assert_eq!(
        run_pascal(
            r#"program T; var parts:array of string; i:Integer; s:string; begin parts:=['a','b','c']; s:=''; for i:=0 to High(parts) do s:=s+parts[i]; WriteLn(s); end."#
        ),
        &["abc"]
    );
}

#[test]
fn sorted_insert_position_linear() {
    assert_eq!(
        run_pascal(
            r#"program T; function PosFor(const a:array of Integer; v:Integer):Integer; var i:Integer; begin Result:=High(a)+1; for i:=Low(a) to High(a) do if a[i]>v then begin Result:=i; Exit; end; end; begin WriteLn(PosFor([1,3,5],4)); end."#
        ),
        &["2"]
    );
}

#[test]
fn generic_bag_add_size() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBag<T>=class private F:array of T; public procedure Add(v:T); function Size:Integer; end; procedure TBag<T>.Add(v:T); var n:Integer; begin n:=Length(F); SetLength(F,n+1); F[n]:=v; end; function TBag<T>.Size:Integer; begin Result:=Length(F); end; var b:TBag<Integer>; begin b:=TBag<Integer>.Create; b.Add(1); b.Add(2); WriteLn(b.Size); b.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn enum_keyed_array_lookup() {
    assert_eq!(
        run_pascal(
            r#"program T; type TK=(A,B,C); var m:array[TK] of Integer; k:TK; begin m[A]:=1; m[B]:=2; m[C]:=3; for k:=A to C do WriteLn(m[k]); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn collection_filter_to_new_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var src,dst:array of Integer; i,n:Integer; begin src:=[1,2,3,4]; n:=0; for i:=0 to High(src) do if src[i] mod 2=0 then begin SetLength(dst,n+1); dst[n]:=src[i]; Inc(n); end; WriteLn(dst[0]); WriteLn(dst[1]); end."#
        ),
        &["2", "4"]
    );
}

#[test]
fn map_default_when_missing() {
    assert_eq!(
        run_pascal(
            r#"program T; function GetDef(const items:array of Integer; idx:Integer; def:Integer):Integer; begin if (idx>=Low(items)) and (idx<=High(items)) then Result:=items[idx] else Result:=def; end; begin WriteLn(GetDef([5],9,-1)); end."#
        ),
        &["-1"]
    );
}

#[test]
fn linked_list_manual_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode=record V:Integer; Next:Integer; end; var nodes:array[0..1] of TNode; begin nodes[0].V:=1; nodes[0].Next:=1; nodes[1].V:=2; nodes[1].Next:=-1; WriteLn(nodes[0].V); WriteLn(nodes[nodes[0].Next].V); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn collection_clear_by_setlength_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of Integer; begin a:=[1,2,3]; SetLength(a,0); WriteLn(Length(a)); end."#
        ),
        &["0"]
    );
}
