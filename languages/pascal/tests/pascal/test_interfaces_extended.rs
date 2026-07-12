/// Extended interface types: reference counting, multiple interfaces, dispatch.
use super::helpers::run_pascal;

#[test]
fn interface_query_returns_same_object() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TObj=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TObj.Create(v:Integer); begin F:=v; end; function TObj.Get:Integer; begin Result:=F; end; var i:IVal; begin i:=TObj.Create(9); WriteLn(i.Get); end."#
        ),
        &["9"]
    );
}

#[test]
fn interface_as_parameter_polymorphism() {
    assert_eq!(
        run_pascal(
            r#"program T; type IShow=interface procedure Show; end; TA=class(TInterfacedObject,IShow) procedure Show; end; TB=class(TInterfacedObject,IShow) procedure Show; end; procedure TA.Show; begin WriteLn('A'); end; procedure TB.Show; begin WriteLn('B'); end; procedure Call(s:IShow); begin s.Show; end; var a:IShow; begin a:=TA.Create; Call(a); a:=TB.Create; Call(a); end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn interface_property_style_getter() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; property Name:string read GetName; end; TPerson=class(TInterfacedObject,IName) private FN:string; public constructor Create(n:string); function GetName:string; end; constructor TPerson.Create(n:string); begin FN:=n; end; function TPerson.GetName:string; begin Result:=FN; end; var p:IName; begin p:=TPerson.Create('Ann'); WriteLn(p.Name); end."#
        ),
        &["Ann"]
    );
}

#[test]
fn two_interfaces_on_one_class() {
    assert_eq!(
        run_pascal(
            r#"program T; type IAdd=interface function Add(a,b:Integer):Integer; end; IMul=interface function Mul(a,b:Integer):Integer; end; TCalc=class(TInterfacedObject,IAdd,IMul) function Add(a,b:Integer):Integer; function Mul(a,b:Integer):Integer; end; function TCalc.Add(a,b:Integer):Integer; begin Result:=a+b; end; function TCalc.Mul(a,b:Integer):Integer; begin Result:=a*b; end; var add:IAdd; mul:IMul; begin add:=TCalc.Create; mul:=TCalc(add); WriteLn(add.Add(2,3)); WriteLn(mul.Mul(2,3)); end."#
        ),
        &["5", "6"]
    );
}

#[test]
fn interface_returned_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T; type IMsg=interface function Text:string; end; TMsg=class(TInterfacedObject,IMsg) private FS:string; public constructor Create(s:string); function Text:string; end; constructor TMsg.Create(s:string); begin FS:=s; end; function TMsg.Text:string; begin Result:=FS; end; function Make(s:string):IMsg; begin Result:=TMsg.Create(s); end; var m:IMsg; begin m:=Make('hi'); WriteLn(m.Text); end."#
        ),
        &["hi"]
    );
}

#[test]
fn interface_nil_before_assign() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; var i:IEmpty; begin if i=nil then WriteLn('nil'); end."#
        ),
        &["nil"]
    );
}

#[test]
fn interface_supports_operator_not_nil() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run; end; TRunner=class(TInterfacedObject,IRun) procedure Run; end; procedure TRunner.Run; begin WriteLn('run'); end; var r:IRun; begin r:=TRunner.Create; if r<>nil then r.Run; end."#
        ),
        &["run"]
    );
}

#[test]
fn interface_array_of_handlers() {
    assert_eq!(
        run_pascal(
            r#"program T; type IHandler=interface procedure Handle(v:Integer); end; TDouble=class(TInterfacedObject,IHandler) procedure Handle(v:Integer); end; procedure TDouble.Handle(v:Integer); begin WriteLn(v*2); end; procedure Dispatch(const hs:array of IHandler; v:Integer); var i:Integer; begin for i:=Low(hs) to High(hs) do hs[i].Handle(v); end; var h:IHandler; begin h:=TDouble.Create; Dispatch([h],3); end."#
        ),
        &["6"]
    );
}

#[test]
fn interface_class_implements_function_result() {
    assert_eq!(
        run_pascal(
            r#"program T; type ILen=interface function GetLen:Integer; end; TStrLen=class(TInterfacedObject,ILen) private FS:string; public constructor Create(s:string); function GetLen:Integer; end; constructor TStrLen.Create(s:string); begin FS:=s; end; function TStrLen.GetLen:Integer; var i:Integer; begin Result:=0; for i:=1 to Length(FS) do Inc(Result); end; function Make(s:string):ILen; begin Result:=TStrLen.Create(s); end; var l:ILen; begin l:=Make('abcd'); WriteLn(l.GetLen); end."#
        ),
        &["4"]
    );
}

#[test]
fn interface_method_chain_two_calls() {
    assert_eq!(
        run_pascal(
            r#"program T; type IInc=interface function Inc(v:Integer):Integer; end; TInc=class(TInterfacedObject,IInc) function Inc(v:Integer):Integer; end; function TInc.Inc(v:Integer):Integer; begin Result:=v+1; end; var i:IInc; begin i:=TInc.Create; WriteLn(i.Inc(i.Inc(1))); end."#
        ),
        &["3"]
    );
}

#[test]
fn interface_stored_in_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type ITag=interface function Tag:Integer; end; TTag=class(TInterfacedObject,ITag) private F:Integer; public constructor Create(t:Integer); function Tag:Integer; end; constructor TTag.Create(t:Integer); begin F:=t; end; function TTag.Tag:Integer; begin Result:=F; end; type TWrap=record H:ITag; end; var w:TWrap; begin w.H:=TTag.Create(7); WriteLn(w.H.Tag); end."#
        ),
        &["7"]
    );
}

#[test]
fn interface_implements_inheritance_extension() {
    assert_eq!(
        run_pascal(
            r#"program T; type IBase=interface function Base:Integer; end; IDerived=interface(IBase) function Derived:Integer; end; TObj=class(TInterfacedObject,IBase,IDerived) function Base:Integer; function Derived:Integer; end; function TObj.Base:Integer; begin Result:=1; end; function TObj.Derived:Integer; begin Result:=2; end; var d:IDerived; begin d:=TObj.Create; WriteLn(d.Base); WriteLn(d.Derived); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn interface_comparer_sort_two_values() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICompare=interface function Less(a,b:Integer):Boolean; end; TAsc=class(TInterfacedObject,ICompare) function Less(a,b:Integer):Boolean; end; function TAsc.Less(a,b:Integer):Boolean; begin Result:=a<b; end; function Min(c:ICompare;a,b:Integer):Integer; begin if c.Less(a,b) then Result:=a else Result:=b; end; var c:ICompare; begin c:=TAsc.Create; WriteLn(Min(c,8,3)); end."#
        ),
        &["3"]
    );
}

#[test]
fn interface_event_callback_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type IClick=interface procedure Click; end; TButton=class(TInterfacedObject,IClick) procedure Click; end; procedure TButton.Click; begin WriteLn('clicked'); end; procedure Wire(evt:IClick); begin evt.Click; end; var b:IClick; begin b:=TButton.Create; Wire(b); end."#
        ),
        &["clicked"]
    );
}

#[test]
fn interface_guid_style_name_only() {
    assert_eq!(
        run_pascal(
            r#"program T; type IWorker=interface ['{11111111-1111-1111-1111-111111111111}'] function Work:Integer; end; TWorker=class(TInterfacedObject,IWorker) function Work:Integer; end; function TWorker.Work:Integer; begin Result:=42; end; var w:IWorker; begin w:=TWorker.Create; WriteLn(w.Work); end."#
        ),
        &["42"]
    );
}

#[test]
fn interface_out_parameter_factory() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function V:Integer; end; TVal=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function V:Integer; end; constructor TVal.Create(v:Integer); begin F:=v; end; function TVal.V:Integer; begin Result:=F; end; procedure Make(out i:IVal); begin i:=TVal.Create(5); end; var x:IVal; begin Make(x); WriteLn(x.V); end."#
        ),
        &["5"]
    );
}

#[test]
fn interface_var_param_updates_holder() {
    assert_eq!(
        run_pascal(
            r#"program T; type ISet=interface procedure Set(v:Integer); end; TSet=class(TInterfacedObject,ISet) private F:Integer; public procedure Set(v:Integer); function Get:Integer; end; procedure TSet.Set(v:Integer); begin F:=v; end; function TSet.Get:Integer; begin Result:=F; end; procedure Apply(s:ISet); begin s.Set(99); end; var o:TSet; begin o:=TSet.Create; Apply(o); WriteLn(o.Get); end."#
        ),
        &["99"]
    );
}

#[test]
fn interface_optional_absent_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; type IMaybe=interface function Has:Boolean; end; procedure Use(m:IMaybe); begin if m=nil then WriteLn('none') else WriteLn('some'); end; begin Use(nil); end."#
        ),
        &["none"]
    );
}

#[test]
fn interface_delegate_to_inner() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICore=interface function Run:Integer; end; TCore=class(TInterfacedObject,ICore) function Run:Integer; end; function TCore.Run:Integer; begin Result:=7; end; type TProxy=class(TInterfacedObject,ICore) private Inner:ICore; public constructor Create(c:ICore); function Run:Integer; end; constructor TProxy.Create(c:ICore); begin Inner:=c; end; function TProxy.Run:Integer; begin Result:=Inner.Run+1; end; var p:ICore; begin p:=TProxy.Create(TCore.Create); WriteLn(p.Run); end."#
        ),
        &["8"]
    );
}

#[test]
fn interface_boolean_result() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICheck=interface function Ok(v:Integer):Boolean; end; TCheck=class(TInterfacedObject,ICheck) function Ok(v:Integer):Boolean; end; function TCheck.Ok(v:Integer):Boolean; begin Result:=v>0; end; var c:ICheck; begin c:=TCheck.Create; WriteLn(c.Ok(1)); WriteLn(c.Ok(-1)); end."#
        ),
        &["true", "false"]
    );
}
