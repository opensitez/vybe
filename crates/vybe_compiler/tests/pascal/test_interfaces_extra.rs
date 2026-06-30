/// Additional interface dispatch and polymorphism scenarios.
use super::helpers::run_pascal;

#[test]
fn iface_extra_add_dispatch_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+1; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(1,1)); end."#
        ),
        &["3"]
    );
}

#[test]
fn iface_extra_param_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item2'); Show(x); end."#
        ),
        &["item2"]
    );
}

#[test]
fn iface_extra_dual_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A3'); end; procedure TBoth.B; begin WriteLn('B3'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A3", "B3"]
    );
}

#[test]
fn iface_extra_factory_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(20); WriteLn(iv.Get); end."#
        ),
        &["20"]
    );
}

#[test]
fn iface_extra_run_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+5); end; var r:IRun; begin r:=TRunner.Create; r.Run(5); end."#
        ),
        &["10"]
    );
}

#[test]
fn iface_extra_empty_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=6; WriteLn(o.Tag); end."#
        ),
        &["6"]
    );
}

#[test]
fn iface_extra_add_dispatch_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+7; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(7,1)); end."#
        ),
        &["15"]
    );
}

#[test]
fn iface_extra_param_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item8'); Show(x); end."#
        ),
        &["item8"]
    );
}

#[test]
fn iface_extra_dual_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A9'); end; procedure TBoth.B; begin WriteLn('B9'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A9", "B9"]
    );
}

#[test]
fn iface_extra_factory_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(50); WriteLn(iv.Get); end."#
        ),
        &["50"]
    );
}

#[test]
fn iface_extra_run_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+11); end; var r:IRun; begin r:=TRunner.Create; r.Run(11); end."#
        ),
        &["22"]
    );
}

#[test]
fn iface_extra_empty_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=12; WriteLn(o.Tag); end."#
        ),
        &["12"]
    );
}

#[test]
fn iface_extra_add_dispatch_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+13; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(13,1)); end."#
        ),
        &["27"]
    );
}

#[test]
fn iface_extra_param_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item14'); Show(x); end."#
        ),
        &["item14"]
    );
}

#[test]
fn iface_extra_dual_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A15'); end; procedure TBoth.B; begin WriteLn('B15'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A15", "B15"]
    );
}

#[test]
fn iface_extra_factory_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(80); WriteLn(iv.Get); end."#
        ),
        &["80"]
    );
}

#[test]
fn iface_extra_run_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+17); end; var r:IRun; begin r:=TRunner.Create; r.Run(17); end."#
        ),
        &["34"]
    );
}

#[test]
fn iface_extra_empty_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=18; WriteLn(o.Tag); end."#
        ),
        &["18"]
    );
}

#[test]
fn iface_extra_add_dispatch_19() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+19; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(19,1)); end."#
        ),
        &["39"]
    );
}

#[test]
fn iface_extra_param_20() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item20'); Show(x); end."#
        ),
        &["item20"]
    );
}

#[test]
fn iface_extra_dual_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A21'); end; procedure TBoth.B; begin WriteLn('B21'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A21", "B21"]
    );
}

#[test]
fn iface_extra_factory_22() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(110); WriteLn(iv.Get); end."#
        ),
        &["110"]
    );
}

#[test]
fn iface_extra_run_23() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+23); end; var r:IRun; begin r:=TRunner.Create; r.Run(23); end."#
        ),
        &["46"]
    );
}

#[test]
fn iface_extra_empty_24() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=24; WriteLn(o.Tag); end."#
        ),
        &["24"]
    );
}

#[test]
fn iface_extra_add_dispatch_25() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+25; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(25,1)); end."#
        ),
        &["51"]
    );
}

#[test]
fn iface_extra_param_26() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item26'); Show(x); end."#
        ),
        &["item26"]
    );
}

#[test]
fn iface_extra_dual_27() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A27'); end; procedure TBoth.B; begin WriteLn('B27'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A27", "B27"]
    );
}

#[test]
fn iface_extra_factory_28() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(140); WriteLn(iv.Get); end."#
        ),
        &["140"]
    );
}

#[test]
fn iface_extra_run_29() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+29); end; var r:IRun; begin r:=TRunner.Create; r.Run(29); end."#
        ),
        &["58"]
    );
}

#[test]
fn iface_extra_empty_30() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=30; WriteLn(o.Tag); end."#
        ),
        &["30"]
    );
}

#[test]
fn iface_extra_add_dispatch_31() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+31; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(31,1)); end."#
        ),
        &["63"]
    );
}

#[test]
fn iface_extra_param_32() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item32'); Show(x); end."#
        ),
        &["item32"]
    );
}

#[test]
fn iface_extra_dual_33() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A33'); end; procedure TBoth.B; begin WriteLn('B33'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A33", "B33"]
    );
}

#[test]
fn iface_extra_factory_34() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(170); WriteLn(iv.Get); end."#
        ),
        &["170"]
    );
}

#[test]
fn iface_extra_run_35() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+35); end; var r:IRun; begin r:=TRunner.Create; r.Run(35); end."#
        ),
        &["70"]
    );
}

#[test]
fn iface_extra_empty_36() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=36; WriteLn(o.Tag); end."#
        ),
        &["36"]
    );
}

#[test]
fn iface_extra_add_dispatch_37() {
    assert_eq!(
        run_pascal(
            r#"program T; type ICalc=interface function Add(a,b:Integer):Integer; end; TImpl=class(TInterfacedObject,ICalc) function Add(a,b:Integer):Integer; end; function TImpl.Add(a,b:Integer):Integer; begin Result:=a+b+37; end; var c:ICalc; begin c:=TImpl.Create; WriteLn(c.Add(37,1)); end."#
        ),
        &["75"]
    );
}

#[test]
fn iface_extra_param_38() {
    assert_eq!(
        run_pascal(
            r#"program T; type IName=interface function GetName:string; end; TNamed=class(TInterfacedObject,IName) private FN:string; public constructor Create(s:string); function GetName:string; end; constructor TNamed.Create(s:string); begin FN:=s; end; function TNamed.GetName:string; begin Result:=FN; end; procedure Show(const n:IName); begin WriteLn(n.GetName); end; var x:IName; begin x:=TNamed.Create('item38'); Show(x); end."#
        ),
        &["item38"]
    );
}

#[test]
fn iface_extra_dual_39() {
    assert_eq!(
        run_pascal(
            r#"program T; type IA=interface procedure A; end; IB=interface procedure B; end; TBoth=class(TInterfacedObject,IA,IB) procedure A; procedure B; end; procedure TBoth.A; begin WriteLn('A39'); end; procedure TBoth.B; begin WriteLn('B39'); end; var a:IA; b:IB; begin a:=TBoth.Create; b:=TBoth(a); a.A; b.B; end."#
        ),
        &["A39", "B39"]
    );
}

#[test]
fn iface_extra_factory_40() {
    assert_eq!(
        run_pascal(
            r#"program T; type IVal=interface function Get:Integer; end; TBox=class(TInterfacedObject,IVal) private F:Integer; public constructor Create(v:Integer); function Get:Integer; end; constructor TBox.Create(v:Integer); begin F:=v; end; function TBox.Get:Integer; begin Result:=F; end; function Make(v:Integer):IVal; begin Result:=TBox.Create(v); end; var iv:IVal; begin iv:=Make(200); WriteLn(iv.Get); end."#
        ),
        &["200"]
    );
}

#[test]
fn iface_extra_run_41() {
    assert_eq!(
        run_pascal(
            r#"program T; type IRun=interface procedure Run(v:Integer); end; TRunner=class(TInterfacedObject,IRun) procedure Run(v:Integer); end; procedure TRunner.Run(v:Integer); begin WriteLn(v+41); end; var r:IRun; begin r:=TRunner.Create; r.Run(41); end."#
        ),
        &["82"]
    );
}

#[test]
fn iface_extra_empty_42() {
    assert_eq!(
        run_pascal(
            r#"program T; type IEmpty=interface end; TObj=class(TInterfacedObject,IEmpty) public Tag:Integer; end; var o:TObj; begin o:=TObj.Create; o.Tag:=42; WriteLn(o.Tag); end."#
        ),
        &["42"]
    );
}
