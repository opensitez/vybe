/// Generic list/map/stack/queue collection patterns.
use super::helpers::run_pascal;

#[test]
fn gcoll_list_add_get_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private F:array of T; public constructor Create; procedure Add(v:T); function Get(i:Integer):T; function Count:Integer; end; constructor TList<T>.Create; begin SetLength(F,0); end; procedure TList<T>.Add(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TList<T>.Get(i:Integer):T; begin Result:=F[i]; end; function TList<T>.Count:Integer; begin Result:=Length(F); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add(1); L.Add(2); WriteLn(L.Get(0)); WriteLn(L.Count); L.Free; end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn gcoll_map_put_get_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMap<K,V>=class private FKeys:array of K; FVals:array of V; public constructor Create; procedure Put(k:K; v:V); function Get(k:K):V; end; constructor TMap<K,V>.Create; begin SetLength(FKeys,0); SetLength(FVals,0); end; procedure TMap<K,V>.Put(k:K; v:V); var l:Integer; begin l:=Length(FKeys); SetLength(FKeys,l+1); SetLength(FVals,l+1); FKeys[l]:=k; FVals[l]:=v; end; function TMap<K,V>.Get(k:K):V; var j:Integer; begin Result:=FVals[0]; for j:=0 to High(FKeys) do if FKeys[j]=k then Result:=FVals[j]; end; var M:TMap<string,Integer>; begin M:=TMap<string,Integer>.Create; M.Put('k2',6); WriteLn(M.Get('k2')); M.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn gcoll_stack_push_pop_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; public constructor Create; procedure Push(v:T); function Pop:T; function Empty:Boolean; end; constructor TStack<T>.Create; begin SetLength(F,0); end; procedure TStack<T>.Push(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TStack<T>.Pop:T; var l:Integer; begin l:=High(F); Result:=F[l]; SetLength(F,l); end; function TStack<T>.Empty:Boolean; begin Result:=Length(F)=0; end; var S:TStack<Integer>; begin S:=TStack<Integer>.Create; S.Push(3); S.Push(4); WriteLn(S.Pop); WriteLn(S.Pop); WriteLn(S.Empty); S.Free; end."#
        ),
        &["4", "3", "true"]
    );
}

#[test]
fn gcoll_queue_fifo_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin SetLength(F,0); FHead:=0; end; procedure TQueue<T>.Enq(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var Q:TQueue<Integer>; begin Q:=TQueue<Integer>.Create; Q.Enq(4); Q.Enq(14); WriteLn(Q.Deq); WriteLn(Q.Deq); Q.Free; end."#
        ),
        &["4", "14"]
    );
}

#[test]
fn gcoll_pair_record_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T,U>=record A:T; B:U; end; var p:TPair<Integer,string>; begin p.A:=5; p.B:='v5'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["5", "v5"]
    );
}

#[test]
fn gcoll_set_contains_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSet<T>=class private F:array of T; public constructor Create; procedure Insert(v:T); function Contains(v:T):Boolean; end; constructor TSet<T>.Create; begin SetLength(F,0); end; procedure TSet<T>.Insert(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TSet<T>.Contains(v:T):Boolean; var j:Integer; begin Result:=false; for j:=0 to High(F) do if F[j]=v then Result:=true; end; var S:TSet<Integer>; begin S:=TSet<Integer>.Create; S.Insert(6); S.Insert(7); WriteLn(S.Contains(6)); WriteLn(S.Contains(105)); S.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn gcoll_linked_node_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=class public Value:T; Next:TNode<T>; constructor Create(v:T); end; constructor TNode<T>.Create(v:T); begin Value:=v; Next:=nil; end; var a,b:TNode<Integer>; begin a:=TNode<Integer>.Create(7); b:=TNode<Integer>.Create(14); a.Next:=b; WriteLn(a.Value); WriteLn(a.Next.Value); a.Free; b.Free; end."#
        ),
        &["7", "14"]
    );
}

#[test]
fn gcoll_list_add_get_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private F:array of T; public constructor Create; procedure Add(v:T); function Get(i:Integer):T; function Count:Integer; end; constructor TList<T>.Create; begin SetLength(F,0); end; procedure TList<T>.Add(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TList<T>.Get(i:Integer):T; begin Result:=F[i]; end; function TList<T>.Count:Integer; begin Result:=Length(F); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add(8); L.Add(16); WriteLn(L.Get(0)); WriteLn(L.Count); L.Free; end."#
        ),
        &["8", "2"]
    );
}

#[test]
fn gcoll_map_put_get_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMap<K,V>=class private FKeys:array of K; FVals:array of V; public constructor Create; procedure Put(k:K; v:V); function Get(k:K):V; end; constructor TMap<K,V>.Create; begin SetLength(FKeys,0); SetLength(FVals,0); end; procedure TMap<K,V>.Put(k:K; v:V); var l:Integer; begin l:=Length(FKeys); SetLength(FKeys,l+1); SetLength(FVals,l+1); FKeys[l]:=k; FVals[l]:=v; end; function TMap<K,V>.Get(k:K):V; var j:Integer; begin Result:=FVals[0]; for j:=0 to High(FKeys) do if FKeys[j]=k then Result:=FVals[j]; end; var M:TMap<string,Integer>; begin M:=TMap<string,Integer>.Create; M.Put('k9',27); WriteLn(M.Get('k9')); M.Free; end."#
        ),
        &["27"]
    );
}

#[test]
fn gcoll_stack_push_pop_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; public constructor Create; procedure Push(v:T); function Pop:T; function Empty:Boolean; end; constructor TStack<T>.Create; begin SetLength(F,0); end; procedure TStack<T>.Push(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TStack<T>.Pop:T; var l:Integer; begin l:=High(F); Result:=F[l]; SetLength(F,l); end; function TStack<T>.Empty:Boolean; begin Result:=Length(F)=0; end; var S:TStack<Integer>; begin S:=TStack<Integer>.Create; S.Push(10); S.Push(11); WriteLn(S.Pop); WriteLn(S.Pop); WriteLn(S.Empty); S.Free; end."#
        ),
        &["11", "10", "true"]
    );
}

#[test]
fn gcoll_queue_fifo_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin SetLength(F,0); FHead:=0; end; procedure TQueue<T>.Enq(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var Q:TQueue<Integer>; begin Q:=TQueue<Integer>.Create; Q.Enq(11); Q.Enq(21); WriteLn(Q.Deq); WriteLn(Q.Deq); Q.Free; end."#
        ),
        &["11", "21"]
    );
}

#[test]
fn gcoll_pair_record_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T,U>=record A:T; B:U; end; var p:TPair<Integer,string>; begin p.A:=12; p.B:='v12'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["12", "v12"]
    );
}

#[test]
fn gcoll_set_contains_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSet<T>=class private F:array of T; public constructor Create; procedure Insert(v:T); function Contains(v:T):Boolean; end; constructor TSet<T>.Create; begin SetLength(F,0); end; procedure TSet<T>.Insert(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TSet<T>.Contains(v:T):Boolean; var j:Integer; begin Result:=false; for j:=0 to High(F) do if F[j]=v then Result:=true; end; var S:TSet<Integer>; begin S:=TSet<Integer>.Create; S.Insert(13); S.Insert(14); WriteLn(S.Contains(13)); WriteLn(S.Contains(112)); S.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn gcoll_linked_node_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=class public Value:T; Next:TNode<T>; constructor Create(v:T); end; constructor TNode<T>.Create(v:T); begin Value:=v; Next:=nil; end; var a,b:TNode<Integer>; begin a:=TNode<Integer>.Create(14); b:=TNode<Integer>.Create(28); a.Next:=b; WriteLn(a.Value); WriteLn(a.Next.Value); a.Free; b.Free; end."#
        ),
        &["14", "28"]
    );
}

#[test]
fn gcoll_list_add_get_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private F:array of T; public constructor Create; procedure Add(v:T); function Get(i:Integer):T; function Count:Integer; end; constructor TList<T>.Create; begin SetLength(F,0); end; procedure TList<T>.Add(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TList<T>.Get(i:Integer):T; begin Result:=F[i]; end; function TList<T>.Count:Integer; begin Result:=Length(F); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add(15); L.Add(30); WriteLn(L.Get(0)); WriteLn(L.Count); L.Free; end."#
        ),
        &["15", "2"]
    );
}

#[test]
fn gcoll_map_put_get_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMap<K,V>=class private FKeys:array of K; FVals:array of V; public constructor Create; procedure Put(k:K; v:V); function Get(k:K):V; end; constructor TMap<K,V>.Create; begin SetLength(FKeys,0); SetLength(FVals,0); end; procedure TMap<K,V>.Put(k:K; v:V); var l:Integer; begin l:=Length(FKeys); SetLength(FKeys,l+1); SetLength(FVals,l+1); FKeys[l]:=k; FVals[l]:=v; end; function TMap<K,V>.Get(k:K):V; var j:Integer; begin Result:=FVals[0]; for j:=0 to High(FKeys) do if FKeys[j]=k then Result:=FVals[j]; end; var M:TMap<string,Integer>; begin M:=TMap<string,Integer>.Create; M.Put('k16',48); WriteLn(M.Get('k16')); M.Free; end."#
        ),
        &["48"]
    );
}

#[test]
fn gcoll_stack_push_pop_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; public constructor Create; procedure Push(v:T); function Pop:T; function Empty:Boolean; end; constructor TStack<T>.Create; begin SetLength(F,0); end; procedure TStack<T>.Push(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TStack<T>.Pop:T; var l:Integer; begin l:=High(F); Result:=F[l]; SetLength(F,l); end; function TStack<T>.Empty:Boolean; begin Result:=Length(F)=0; end; var S:TStack<Integer>; begin S:=TStack<Integer>.Create; S.Push(17); S.Push(18); WriteLn(S.Pop); WriteLn(S.Pop); WriteLn(S.Empty); S.Free; end."#
        ),
        &["18", "17", "true"]
    );
}

#[test]
fn gcoll_queue_fifo_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin SetLength(F,0); FHead:=0; end; procedure TQueue<T>.Enq(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var Q:TQueue<Integer>; begin Q:=TQueue<Integer>.Create; Q.Enq(18); Q.Enq(28); WriteLn(Q.Deq); WriteLn(Q.Deq); Q.Free; end."#
        ),
        &["18", "28"]
    );
}

#[test]
fn gcoll_pair_record_19() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T,U>=record A:T; B:U; end; var p:TPair<Integer,string>; begin p.A:=19; p.B:='v19'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["19", "v19"]
    );
}

#[test]
fn gcoll_set_contains_20() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSet<T>=class private F:array of T; public constructor Create; procedure Insert(v:T); function Contains(v:T):Boolean; end; constructor TSet<T>.Create; begin SetLength(F,0); end; procedure TSet<T>.Insert(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TSet<T>.Contains(v:T):Boolean; var j:Integer; begin Result:=false; for j:=0 to High(F) do if F[j]=v then Result:=true; end; var S:TSet<Integer>; begin S:=TSet<Integer>.Create; S.Insert(20); S.Insert(21); WriteLn(S.Contains(20)); WriteLn(S.Contains(119)); S.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn gcoll_linked_node_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=class public Value:T; Next:TNode<T>; constructor Create(v:T); end; constructor TNode<T>.Create(v:T); begin Value:=v; Next:=nil; end; var a,b:TNode<Integer>; begin a:=TNode<Integer>.Create(21); b:=TNode<Integer>.Create(42); a.Next:=b; WriteLn(a.Value); WriteLn(a.Next.Value); a.Free; b.Free; end."#
        ),
        &["21", "42"]
    );
}

#[test]
fn gcoll_list_add_get_22() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private F:array of T; public constructor Create; procedure Add(v:T); function Get(i:Integer):T; function Count:Integer; end; constructor TList<T>.Create; begin SetLength(F,0); end; procedure TList<T>.Add(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TList<T>.Get(i:Integer):T; begin Result:=F[i]; end; function TList<T>.Count:Integer; begin Result:=Length(F); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add(22); L.Add(44); WriteLn(L.Get(0)); WriteLn(L.Count); L.Free; end."#
        ),
        &["22", "2"]
    );
}

#[test]
fn gcoll_map_put_get_23() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMap<K,V>=class private FKeys:array of K; FVals:array of V; public constructor Create; procedure Put(k:K; v:V); function Get(k:K):V; end; constructor TMap<K,V>.Create; begin SetLength(FKeys,0); SetLength(FVals,0); end; procedure TMap<K,V>.Put(k:K; v:V); var l:Integer; begin l:=Length(FKeys); SetLength(FKeys,l+1); SetLength(FVals,l+1); FKeys[l]:=k; FVals[l]:=v; end; function TMap<K,V>.Get(k:K):V; var j:Integer; begin Result:=FVals[0]; for j:=0 to High(FKeys) do if FKeys[j]=k then Result:=FVals[j]; end; var M:TMap<string,Integer>; begin M:=TMap<string,Integer>.Create; M.Put('k23',69); WriteLn(M.Get('k23')); M.Free; end."#
        ),
        &["69"]
    );
}

#[test]
fn gcoll_stack_push_pop_24() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; public constructor Create; procedure Push(v:T); function Pop:T; function Empty:Boolean; end; constructor TStack<T>.Create; begin SetLength(F,0); end; procedure TStack<T>.Push(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TStack<T>.Pop:T; var l:Integer; begin l:=High(F); Result:=F[l]; SetLength(F,l); end; function TStack<T>.Empty:Boolean; begin Result:=Length(F)=0; end; var S:TStack<Integer>; begin S:=TStack<Integer>.Create; S.Push(24); S.Push(25); WriteLn(S.Pop); WriteLn(S.Pop); WriteLn(S.Empty); S.Free; end."#
        ),
        &["25", "24", "true"]
    );
}

#[test]
fn gcoll_queue_fifo_25() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin SetLength(F,0); FHead:=0; end; procedure TQueue<T>.Enq(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var Q:TQueue<Integer>; begin Q:=TQueue<Integer>.Create; Q.Enq(25); Q.Enq(35); WriteLn(Q.Deq); WriteLn(Q.Deq); Q.Free; end."#
        ),
        &["25", "35"]
    );
}

#[test]
fn gcoll_pair_record_26() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T,U>=record A:T; B:U; end; var p:TPair<Integer,string>; begin p.A:=26; p.B:='v26'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["26", "v26"]
    );
}

#[test]
fn gcoll_set_contains_27() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSet<T>=class private F:array of T; public constructor Create; procedure Insert(v:T); function Contains(v:T):Boolean; end; constructor TSet<T>.Create; begin SetLength(F,0); end; procedure TSet<T>.Insert(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TSet<T>.Contains(v:T):Boolean; var j:Integer; begin Result:=false; for j:=0 to High(F) do if F[j]=v then Result:=true; end; var S:TSet<Integer>; begin S:=TSet<Integer>.Create; S.Insert(27); S.Insert(28); WriteLn(S.Contains(27)); WriteLn(S.Contains(126)); S.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn gcoll_linked_node_28() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=class public Value:T; Next:TNode<T>; constructor Create(v:T); end; constructor TNode<T>.Create(v:T); begin Value:=v; Next:=nil; end; var a,b:TNode<Integer>; begin a:=TNode<Integer>.Create(28); b:=TNode<Integer>.Create(56); a.Next:=b; WriteLn(a.Value); WriteLn(a.Next.Value); a.Free; b.Free; end."#
        ),
        &["28", "56"]
    );
}

#[test]
fn gcoll_list_add_get_29() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private F:array of T; public constructor Create; procedure Add(v:T); function Get(i:Integer):T; function Count:Integer; end; constructor TList<T>.Create; begin SetLength(F,0); end; procedure TList<T>.Add(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TList<T>.Get(i:Integer):T; begin Result:=F[i]; end; function TList<T>.Count:Integer; begin Result:=Length(F); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add(29); L.Add(58); WriteLn(L.Get(0)); WriteLn(L.Count); L.Free; end."#
        ),
        &["29", "2"]
    );
}

#[test]
fn gcoll_map_put_get_30() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMap<K,V>=class private FKeys:array of K; FVals:array of V; public constructor Create; procedure Put(k:K; v:V); function Get(k:K):V; end; constructor TMap<K,V>.Create; begin SetLength(FKeys,0); SetLength(FVals,0); end; procedure TMap<K,V>.Put(k:K; v:V); var l:Integer; begin l:=Length(FKeys); SetLength(FKeys,l+1); SetLength(FVals,l+1); FKeys[l]:=k; FVals[l]:=v; end; function TMap<K,V>.Get(k:K):V; var j:Integer; begin Result:=FVals[0]; for j:=0 to High(FKeys) do if FKeys[j]=k then Result:=FVals[j]; end; var M:TMap<string,Integer>; begin M:=TMap<string,Integer>.Create; M.Put('k30',90); WriteLn(M.Get('k30')); M.Free; end."#
        ),
        &["90"]
    );
}

#[test]
fn gcoll_stack_push_pop_31() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; public constructor Create; procedure Push(v:T); function Pop:T; function Empty:Boolean; end; constructor TStack<T>.Create; begin SetLength(F,0); end; procedure TStack<T>.Push(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TStack<T>.Pop:T; var l:Integer; begin l:=High(F); Result:=F[l]; SetLength(F,l); end; function TStack<T>.Empty:Boolean; begin Result:=Length(F)=0; end; var S:TStack<Integer>; begin S:=TStack<Integer>.Create; S.Push(31); S.Push(32); WriteLn(S.Pop); WriteLn(S.Pop); WriteLn(S.Empty); S.Free; end."#
        ),
        &["32", "31", "true"]
    );
}

#[test]
fn gcoll_queue_fifo_32() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin SetLength(F,0); FHead:=0; end; procedure TQueue<T>.Enq(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var Q:TQueue<Integer>; begin Q:=TQueue<Integer>.Create; Q.Enq(32); Q.Enq(42); WriteLn(Q.Deq); WriteLn(Q.Deq); Q.Free; end."#
        ),
        &["32", "42"]
    );
}

#[test]
fn gcoll_pair_record_33() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T,U>=record A:T; B:U; end; var p:TPair<Integer,string>; begin p.A:=33; p.B:='v33'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["33", "v33"]
    );
}

#[test]
fn gcoll_set_contains_34() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSet<T>=class private F:array of T; public constructor Create; procedure Insert(v:T); function Contains(v:T):Boolean; end; constructor TSet<T>.Create; begin SetLength(F,0); end; procedure TSet<T>.Insert(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TSet<T>.Contains(v:T):Boolean; var j:Integer; begin Result:=false; for j:=0 to High(F) do if F[j]=v then Result:=true; end; var S:TSet<Integer>; begin S:=TSet<Integer>.Create; S.Insert(34); S.Insert(35); WriteLn(S.Contains(34)); WriteLn(S.Contains(133)); S.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn gcoll_linked_node_35() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=class public Value:T; Next:TNode<T>; constructor Create(v:T); end; constructor TNode<T>.Create(v:T); begin Value:=v; Next:=nil; end; var a,b:TNode<Integer>; begin a:=TNode<Integer>.Create(35); b:=TNode<Integer>.Create(70); a.Next:=b; WriteLn(a.Value); WriteLn(a.Next.Value); a.Free; b.Free; end."#
        ),
        &["35", "70"]
    );
}

#[test]
fn gcoll_list_add_get_36() {
    assert_eq!(
        run_pascal(
            r#"program T; type TList<T>=class private F:array of T; public constructor Create; procedure Add(v:T); function Get(i:Integer):T; function Count:Integer; end; constructor TList<T>.Create; begin SetLength(F,0); end; procedure TList<T>.Add(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TList<T>.Get(i:Integer):T; begin Result:=F[i]; end; function TList<T>.Count:Integer; begin Result:=Length(F); end; var L:TList<Integer>; begin L:=TList<Integer>.Create; L.Add(36); L.Add(72); WriteLn(L.Get(0)); WriteLn(L.Count); L.Free; end."#
        ),
        &["36", "2"]
    );
}

#[test]
fn gcoll_map_put_get_37() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMap<K,V>=class private FKeys:array of K; FVals:array of V; public constructor Create; procedure Put(k:K; v:V); function Get(k:K):V; end; constructor TMap<K,V>.Create; begin SetLength(FKeys,0); SetLength(FVals,0); end; procedure TMap<K,V>.Put(k:K; v:V); var l:Integer; begin l:=Length(FKeys); SetLength(FKeys,l+1); SetLength(FVals,l+1); FKeys[l]:=k; FVals[l]:=v; end; function TMap<K,V>.Get(k:K):V; var j:Integer; begin Result:=FVals[0]; for j:=0 to High(FKeys) do if FKeys[j]=k then Result:=FVals[j]; end; var M:TMap<string,Integer>; begin M:=TMap<string,Integer>.Create; M.Put('k37',111); WriteLn(M.Get('k37')); M.Free; end."#
        ),
        &["111"]
    );
}

#[test]
fn gcoll_stack_push_pop_38() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStack<T>=class private F:array of T; public constructor Create; procedure Push(v:T); function Pop:T; function Empty:Boolean; end; constructor TStack<T>.Create; begin SetLength(F,0); end; procedure TStack<T>.Push(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TStack<T>.Pop:T; var l:Integer; begin l:=High(F); Result:=F[l]; SetLength(F,l); end; function TStack<T>.Empty:Boolean; begin Result:=Length(F)=0; end; var S:TStack<Integer>; begin S:=TStack<Integer>.Create; S.Push(38); S.Push(39); WriteLn(S.Pop); WriteLn(S.Pop); WriteLn(S.Empty); S.Free; end."#
        ),
        &["39", "38", "true"]
    );
}

#[test]
fn gcoll_queue_fifo_39() {
    assert_eq!(
        run_pascal(
            r#"program T; type TQueue<T>=class private F:array of T; FHead:Integer; public constructor Create; procedure Enq(v:T); function Deq:T; end; constructor TQueue<T>.Create; begin SetLength(F,0); FHead:=0; end; procedure TQueue<T>.Enq(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TQueue<T>.Deq:T; begin Result:=F[FHead]; Inc(FHead); end; var Q:TQueue<Integer>; begin Q:=TQueue<Integer>.Create; Q.Enq(39); Q.Enq(49); WriteLn(Q.Deq); WriteLn(Q.Deq); Q.Free; end."#
        ),
        &["39", "49"]
    );
}

#[test]
fn gcoll_pair_record_40() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair<T,U>=record A:T; B:U; end; var p:TPair<Integer,string>; begin p.A:=40; p.B:='v40'; WriteLn(p.A); WriteLn(p.B); end."#
        ),
        &["40", "v40"]
    );
}

#[test]
fn gcoll_set_contains_41() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSet<T>=class private F:array of T; public constructor Create; procedure Insert(v:T); function Contains(v:T):Boolean; end; constructor TSet<T>.Create; begin SetLength(F,0); end; procedure TSet<T>.Insert(v:T); var l:Integer; begin l:=Length(F); SetLength(F,l+1); F[l]:=v; end; function TSet<T>.Contains(v:T):Boolean; var j:Integer; begin Result:=false; for j:=0 to High(F) do if F[j]=v then Result:=true; end; var S:TSet<Integer>; begin S:=TSet<Integer>.Create; S.Insert(41); S.Insert(42); WriteLn(S.Contains(41)); WriteLn(S.Contains(140)); S.Free; end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn gcoll_linked_node_42() {
    assert_eq!(
        run_pascal(
            r#"program T; type TNode<T>=class public Value:T; Next:TNode<T>; constructor Create(v:T); end; constructor TNode<T>.Create(v:T); begin Value:=v; Next:=nil; end; var a,b:TNode<Integer>; begin a:=TNode<Integer>.Create(42); b:=TNode<Integer>.Create(84); a.Next:=b; WriteLn(a.Value); WriteLn(a.Next.Value); a.Free; b.Free; end."#
        ),
        &["42", "84"]
    );
}
