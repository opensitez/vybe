/// Dynamic string arrays and string list operations.
use super::helpers::run_pascal;

#[test]
fn dstrarr_two_elem_1() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a1'; a[1]:='b1'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a1", "b1"]
    );
}

#[test]
fn dstrarr_join_2() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x2','y2']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x2y2"]
    );
}

#[test]
fn dstrarr_length_3() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only3'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only3"]
    );
}

#[test]
fn dstrarr_append_4() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m4']; SetLength(a,2); a[1]:='n4'; WriteLn(a[1]); end."#
        ),
        &["n4"]
    );
}

#[test]
fn dstrarr_find_5() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b5','c']; c:=0; for i:=0 to High(a) do if a[i]='b5' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_6() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,2); a[0]:='z6'; WriteLn(a[0]); end."#
        ),
        &["z6"]
    );
}

#[test]
fn dstrarr_two_elem_7() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a7'; a[1]:='b7'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a7", "b7"]
    );
}

#[test]
fn dstrarr_join_8() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x8','y8']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x8y8"]
    );
}

#[test]
fn dstrarr_length_9() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only9'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only9"]
    );
}

#[test]
fn dstrarr_append_10() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m10']; SetLength(a,2); a[1]:='n10'; WriteLn(a[1]); end."#
        ),
        &["n10"]
    );
}

#[test]
fn dstrarr_find_11() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b11','c']; c:=0; for i:=0 to High(a) do if a[i]='b11' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_12() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,3); a[0]:='z12'; WriteLn(a[0]); end."#
        ),
        &["z12"]
    );
}

#[test]
fn dstrarr_two_elem_13() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a13'; a[1]:='b13'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a13", "b13"]
    );
}

#[test]
fn dstrarr_join_14() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x14','y14']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x14y14"]
    );
}

#[test]
fn dstrarr_length_15() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only15'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only15"]
    );
}

#[test]
fn dstrarr_append_16() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m16']; SetLength(a,2); a[1]:='n16'; WriteLn(a[1]); end."#
        ),
        &["n16"]
    );
}

#[test]
fn dstrarr_find_17() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b17','c']; c:=0; for i:=0 to High(a) do if a[i]='b17' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_18() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,4); a[0]:='z18'; WriteLn(a[0]); end."#
        ),
        &["z18"]
    );
}

#[test]
fn dstrarr_two_elem_19() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a19'; a[1]:='b19'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a19", "b19"]
    );
}

#[test]
fn dstrarr_join_20() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x20','y20']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x20y20"]
    );
}

#[test]
fn dstrarr_length_21() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only21'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only21"]
    );
}

#[test]
fn dstrarr_append_22() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m22']; SetLength(a,2); a[1]:='n22'; WriteLn(a[1]); end."#
        ),
        &["n22"]
    );
}

#[test]
fn dstrarr_find_23() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b23','c']; c:=0; for i:=0 to High(a) do if a[i]='b23' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_24() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,5); a[0]:='z24'; WriteLn(a[0]); end."#
        ),
        &["z24"]
    );
}

#[test]
fn dstrarr_two_elem_25() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a25'; a[1]:='b25'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a25", "b25"]
    );
}

#[test]
fn dstrarr_join_26() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x26','y26']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x26y26"]
    );
}

#[test]
fn dstrarr_length_27() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only27'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only27"]
    );
}

#[test]
fn dstrarr_append_28() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m28']; SetLength(a,2); a[1]:='n28'; WriteLn(a[1]); end."#
        ),
        &["n28"]
    );
}

#[test]
fn dstrarr_find_29() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b29','c']; c:=0; for i:=0 to High(a) do if a[i]='b29' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_30() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,1); a[0]:='z30'; WriteLn(a[0]); end."#
        ),
        &["z30"]
    );
}

#[test]
fn dstrarr_two_elem_31() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a31'; a[1]:='b31'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a31", "b31"]
    );
}

#[test]
fn dstrarr_join_32() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x32','y32']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x32y32"]
    );
}

#[test]
fn dstrarr_length_33() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only33'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only33"]
    );
}

#[test]
fn dstrarr_append_34() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m34']; SetLength(a,2); a[1]:='n34'; WriteLn(a[1]); end."#
        ),
        &["n34"]
    );
}

#[test]
fn dstrarr_find_35() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b35','c']; c:=0; for i:=0 to High(a) do if a[i]='b35' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_36() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,2); a[0]:='z36'; WriteLn(a[0]); end."#
        ),
        &["z36"]
    );
}

#[test]
fn dstrarr_two_elem_37() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,2); a[0]:='a37'; a[1]:='b37'; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["a37", "b37"]
    );
}

#[test]
fn dstrarr_join_38() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i:Integer; s:string; begin a:=['x38','y38']; s:=''; for i:=0 to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["x38y38"]
    );
}

#[test]
fn dstrarr_length_39() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,1); a[0]:='only39'; WriteLn(Length(a)); WriteLn(a[0]); end."#
        ),
        &["1", "only39"]
    );
}

#[test]
fn dstrarr_append_40() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['m40']; SetLength(a,2); a[1]:='n40'; WriteLn(a[1]); end."#
        ),
        &["n40"]
    );
}

#[test]
fn dstrarr_find_41() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; i,c:Integer; begin a:=['a','b41','c']; c:=0; for i:=0 to High(a) do if a[i]='b41' then Inc(c); WriteLn(c); end."#
        ),
        &["1"]
    );
}

#[test]
fn dstrarr_resize_42() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin SetLength(a,0); SetLength(a,3); a[0]:='z42'; WriteLn(a[0]); end."#
        ),
        &["z42"]
    );
}
