/// Property read/write with getters and setters.
use super::helpers::run_pascal;

#[test]
fn propacc_direct_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=1; WriteLn(o.Val); o.Free; end."#
        ),
        &["1"]
    );
}

#[test]
fn propacc_getter_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+2; end; var o:T; begin o:=T.Create; o.F:=2; WriteLn(o.Val); o.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn propacc_setter_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+3; end; var o:T; begin o:=T.Create; o.Val:=3; WriteLn(o.F); o.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn propacc_two_props_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=4; o.Y:=5; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn propacc_rw_custom_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=10; WriteLn(o.Double); o.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn propacc_string_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_6'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_6"]
    );
}

#[test]
fn propacc_direct_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=7; WriteLn(o.Val); o.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn propacc_getter_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+8; end; var o:T; begin o:=T.Create; o.F:=8; WriteLn(o.Val); o.Free; end."#
        ),
        &["16"]
    );
}

#[test]
fn propacc_setter_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+9; end; var o:T; begin o:=T.Create; o.Val:=9; WriteLn(o.F); o.Free; end."#
        ),
        &["18"]
    );
}

#[test]
fn propacc_two_props_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=10; o.Y:=11; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["21"]
    );
}

#[test]
fn propacc_rw_custom_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=22; WriteLn(o.Double); o.Free; end."#
        ),
        &["22"]
    );
}

#[test]
fn propacc_string_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_12'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_12"]
    );
}

#[test]
fn propacc_direct_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=13; WriteLn(o.Val); o.Free; end."#
        ),
        &["13"]
    );
}

#[test]
fn propacc_getter_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+14; end; var o:T; begin o:=T.Create; o.F:=14; WriteLn(o.Val); o.Free; end."#
        ),
        &["28"]
    );
}

#[test]
fn propacc_setter_15() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+15; end; var o:T; begin o:=T.Create; o.Val:=15; WriteLn(o.F); o.Free; end."#
        ),
        &["30"]
    );
}

#[test]
fn propacc_two_props_16() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=16; o.Y:=17; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["33"]
    );
}

#[test]
fn propacc_rw_custom_17() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=34; WriteLn(o.Double); o.Free; end."#
        ),
        &["34"]
    );
}

#[test]
fn propacc_string_18() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_18'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_18"]
    );
}

#[test]
fn propacc_direct_19() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=19; WriteLn(o.Val); o.Free; end."#
        ),
        &["19"]
    );
}

#[test]
fn propacc_getter_20() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+20; end; var o:T; begin o:=T.Create; o.F:=20; WriteLn(o.Val); o.Free; end."#
        ),
        &["40"]
    );
}

#[test]
fn propacc_setter_21() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+21; end; var o:T; begin o:=T.Create; o.Val:=21; WriteLn(o.F); o.Free; end."#
        ),
        &["42"]
    );
}

#[test]
fn propacc_two_props_22() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=22; o.Y:=23; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["45"]
    );
}

#[test]
fn propacc_rw_custom_23() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=46; WriteLn(o.Double); o.Free; end."#
        ),
        &["46"]
    );
}

#[test]
fn propacc_string_24() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_24'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_24"]
    );
}

#[test]
fn propacc_direct_25() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=25; WriteLn(o.Val); o.Free; end."#
        ),
        &["25"]
    );
}

#[test]
fn propacc_getter_26() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+26; end; var o:T; begin o:=T.Create; o.F:=26; WriteLn(o.Val); o.Free; end."#
        ),
        &["52"]
    );
}

#[test]
fn propacc_setter_27() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+27; end; var o:T; begin o:=T.Create; o.Val:=27; WriteLn(o.F); o.Free; end."#
        ),
        &["54"]
    );
}

#[test]
fn propacc_two_props_28() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=28; o.Y:=29; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["57"]
    );
}

#[test]
fn propacc_rw_custom_29() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=58; WriteLn(o.Double); o.Free; end."#
        ),
        &["58"]
    );
}

#[test]
fn propacc_string_30() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_30'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_30"]
    );
}

#[test]
fn propacc_direct_31() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=31; WriteLn(o.Val); o.Free; end."#
        ),
        &["31"]
    );
}

#[test]
fn propacc_getter_32() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+32; end; var o:T; begin o:=T.Create; o.F:=32; WriteLn(o.Val); o.Free; end."#
        ),
        &["64"]
    );
}

#[test]
fn propacc_setter_33() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+33; end; var o:T; begin o:=T.Create; o.Val:=33; WriteLn(o.F); o.Free; end."#
        ),
        &["66"]
    );
}

#[test]
fn propacc_two_props_34() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=34; o.Y:=35; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["69"]
    );
}

#[test]
fn propacc_rw_custom_35() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=70; WriteLn(o.Double); o.Free; end."#
        ),
        &["70"]
    );
}

#[test]
fn propacc_string_36() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_36'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_36"]
    );
}

#[test]
fn propacc_direct_37() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; public property Val:Integer read F write F; end; var o:T; begin o:=T.Create; o.Val:=37; WriteLn(o.Val); o.Free; end."#
        ),
        &["37"]
    );
}

#[test]
fn propacc_getter_38() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetV:Integer; public property Val:Integer read GetV; end; function T.GetV:Integer; begin Result:=F+38; end; var o:T; begin o:=T.Create; o.F:=38; WriteLn(o.Val); o.Free; end."#
        ),
        &["76"]
    );
}

#[test]
fn propacc_setter_39() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; procedure SetV(v:Integer); public property Val:Integer read F write SetV; end; procedure T.SetV(v:Integer); begin F:=v+39; end; var o:T; begin o:=T.Create; o.Val:=39; WriteLn(o.F); o.Free; end."#
        ),
        &["78"]
    );
}

#[test]
fn propacc_two_props_40() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var o:T; begin o:=T.Create; o.X:=40; o.Y:=41; WriteLn(o.X+o.Y); o.Free; end."#
        ),
        &["81"]
    );
}

#[test]
fn propacc_rw_custom_41() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private F:Integer; function GetD:Integer; procedure SetD(v:Integer); public property Double:Integer read GetD write SetD; end; function T.GetD:Integer; begin Result:=F*2; end; procedure T.SetD(v:Integer); begin F:=v div 2; end; var o:T; begin o:=T.Create; o.Double:=82; WriteLn(o.Double); o.Free; end."#
        ),
        &["82"]
    );
}

#[test]
fn propacc_string_42() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=class private FName:string; function GetName:string; procedure SetName(const s:string); public property Name:string read GetName write SetName; end; function T.GetName:string; begin Result:=FName; end; procedure T.SetName(const s:string); begin FName:=s+'_42'; end; var o:T; begin o:=T.Create; o.Name:='x'; WriteLn(o.Name); o.Free; end."#
        ),
        &["x_42"]
    );
}
