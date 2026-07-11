/// Complex record compositions, nesting, and variant parts.
use super::helpers::run_pascal;

#[test]
fn nested_record_1_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=1; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["1"]
    );
}

#[test]
fn nested_record_2_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=2; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_record_3_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=3; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["3"]
    );
}

#[test]
fn nested_record_4_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=4; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["4"]
    );
}

#[test]
fn nested_record_5_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=5; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_record_6_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=6; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["6"]
    );
}

#[test]
fn nested_record_7_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=7; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_record_8_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=8; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["8"]
    );
}

#[test]
fn nested_record_9_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=9; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["9"]
    );
}

#[test]
fn nested_record_10_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=10; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["10"]
    );
}

#[test]
fn nested_record_11_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=11; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["11"]
    );
}

#[test]
fn nested_record_12_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=12; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_record_13_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=13; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["13"]
    );
}

#[test]
fn nested_record_14_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; type T1=record V:Integer; end; type T2=record Inner:T1; end; type T3=record Mid:T2; end; var o:T3; begin o.Mid.Inner.V:=14; WriteLn(o.Mid.Inner.V); end."#
        ),
        &["14"]
    );
}

#[test]
fn case_record_int_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=2; WriteLn(v.I); end."#
        ),
        &["2"]
    );
}

#[test]
fn case_record_int_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=4; WriteLn(v.I); end."#
        ),
        &["4"]
    );
}

#[test]
fn case_record_int_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=6; WriteLn(v.I); end."#
        ),
        &["6"]
    );
}

#[test]
fn case_record_int_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=8; WriteLn(v.I); end."#
        ),
        &["8"]
    );
}

#[test]
fn case_record_int_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=10; WriteLn(v.I); end."#
        ),
        &["10"]
    );
}

#[test]
fn case_record_int_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=12; WriteLn(v.I); end."#
        ),
        &["12"]
    );
}

#[test]
fn case_record_int_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=14; WriteLn(v.I); end."#
        ),
        &["14"]
    );
}

#[test]
fn case_record_int_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=16; WriteLn(v.I); end."#
        ),
        &["16"]
    );
}

#[test]
fn case_record_int_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=18; WriteLn(v.I); end."#
        ),
        &["18"]
    );
}

#[test]
fn case_record_int_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=20; WriteLn(v.I); end."#
        ),
        &["20"]
    );
}

#[test]
fn case_record_int_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=22; WriteLn(v.I); end."#
        ),
        &["22"]
    );
}

#[test]
fn case_record_int_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=24; WriteLn(v.I); end."#
        ),
        &["24"]
    );
}

#[test]
fn case_record_int_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=26; WriteLn(v.I); end."#
        ),
        &["26"]
    );
}

#[test]
fn case_record_int_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TVal=record case K:Integer of 0:(I:Integer); 1:(S:string); end; var v:TVal; begin v.K:=0; v.I:=28; WriteLn(v.I); end."#
        ),
        &["28"]
    );
}

#[test]
fn record_with_method_1() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=1; p.B:=2; WriteLn(p.Sum); end."#
        ),
        &["3"]
    );
}

#[test]
fn record_with_method_2() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=2; p.B:=3; WriteLn(p.Sum); end."#
        ),
        &["5"]
    );
}

#[test]
fn record_with_method_3() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=3; p.B:=4; WriteLn(p.Sum); end."#
        ),
        &["7"]
    );
}

#[test]
fn record_with_method_4() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=4; p.B:=5; WriteLn(p.Sum); end."#
        ),
        &["9"]
    );
}

#[test]
fn record_with_method_5() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=5; p.B:=6; WriteLn(p.Sum); end."#
        ),
        &["11"]
    );
}

#[test]
fn record_with_method_6() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=6; p.B:=7; WriteLn(p.Sum); end."#
        ),
        &["13"]
    );
}

#[test]
fn record_with_method_7() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=7; p.B:=8; WriteLn(p.Sum); end."#
        ),
        &["15"]
    );
}

#[test]
fn record_with_method_8() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=8; p.B:=9; WriteLn(p.Sum); end."#
        ),
        &["17"]
    );
}

#[test]
fn record_with_method_9() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=9; p.B:=10; WriteLn(p.Sum); end."#
        ),
        &["19"]
    );
}

#[test]
fn record_with_method_10() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=10; p.B:=11; WriteLn(p.Sum); end."#
        ),
        &["21"]
    );
}

#[test]
fn record_with_method_11() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=11; p.B:=12; WriteLn(p.Sum); end."#
        ),
        &["23"]
    );
}

#[test]
fn record_with_method_12() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=12; p.B:=13; WriteLn(p.Sum); end."#
        ),
        &["25"]
    );
}

#[test]
fn record_with_method_13() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=13; p.B:=14; WriteLn(p.Sum); end."#
        ),
        &["27"]
    );
}

#[test]
fn record_with_method_14() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPair=record A,B:Integer; function Sum:Integer; end; function TPair.Sum:Integer; begin Result:=A+B; end; var p:TPair; begin p.A:=14; p.B:=15; WriteLn(p.Sum); end."#
        ),
        &["29"]
    );
}
