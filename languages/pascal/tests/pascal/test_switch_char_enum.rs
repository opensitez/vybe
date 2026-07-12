/// Case statements on Char values and enumeration types.
use super::helpers::run_pascal;

#[test]
fn case_char_single_a() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='a'; case c of 'a': WriteLn('alpha'); 'b': WriteLn('beta'); else WriteLn('?'); end; end."#
        ),
        &["alpha"]
    );
}

#[test]
fn case_char_single_z() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='z'; case c of 'a': WriteLn('a'); 'z': WriteLn('z'); else WriteLn('?'); end; end."#
        ),
        &["z"]
    );
}

#[test]
fn case_char_range_lower_half() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='f'; case c of 'a'..'m': WriteLn('first'); 'n'..'z': WriteLn('second'); end; end."#
        ),
        &["first"]
    );
}

#[test]
fn case_char_range_upper_half() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='t'; case c of 'a'..'m': WriteLn('first'); 'n'..'z': WriteLn('second'); end; end."#
        ),
        &["second"]
    );
}

#[test]
fn case_char_digit() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='7'; case c of '0'..'9': WriteLn('digit'); else WriteLn('other'); end; end."#
        ),
        &["digit"]
    );
}

#[test]
fn case_char_space() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:=' '; case c of ' ': WriteLn('space'); else WriteLn('n'); end; end."#
        ),
        &["space"]
    );
}

#[test]
fn case_char_comma_list() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='x'; case c of 'a','e','i','o','u': WriteLn('vowel'); else WriteLn('cons'); end; end."#
        ),
        &["cons"]
    );
}

#[test]
fn case_char_vowel_e() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='e'; case c of 'a','e','i','o','u': WriteLn('vowel'); else WriteLn('cons'); end; end."#
        ),
        &["vowel"]
    );
}

#[test]
fn case_char_upper_a() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='A'; case c of 'A'..'Z': WriteLn('upper'); 'a'..'z': WriteLn('lower'); end; end."#
        ),
        &["upper"]
    );
}

#[test]
fn case_char_lower_x() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='x'; case c of 'A'..'Z': WriteLn('upper'); 'a'..'z': WriteLn('lower'); end; end."#
        ),
        &["lower"]
    );
}

#[test]
fn case_char_punctuation() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='!'; case c of '!','?': WriteLn('punct'); else WriteLn('plain'); end; end."#
        ),
        &["punct"]
    );
}

#[test]
fn case_char_tab_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:=#9; case c of 'a'..'z': WriteLn('letter'); else WriteLn('other'); end; end."#
        ),
        &["other"]
    );
}

#[test]
fn case_enum_red() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor=(Red,Green,Blue); var c:TColor; begin c:=Red; case c of Red:WriteLn('r'); Green:WriteLn('g'); Blue:WriteLn('b'); end; end."#
        ),
        &["r"]
    );
}

#[test]
fn case_enum_green() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor=(Red,Green,Blue); var c:TColor; begin c:=Green; case c of Red:WriteLn('r'); Green:WriteLn('g'); Blue:WriteLn('b'); end; end."#
        ),
        &["g"]
    );
}

#[test]
fn case_enum_blue() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor=(Red,Green,Blue); var c:TColor; begin c:=Blue; case c of Red:WriteLn('r'); Green:WriteLn('g'); Blue:WriteLn('b'); end; end."#
        ),
        &["b"]
    );
}

#[test]
fn case_enum_else_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDir=(N,E,S,W); var d:TDir; begin d:=W; case d of N:WriteLn('n'); E:WriteLn('e'); else WriteLn('other'); end; end."#
        ),
        &["other"]
    );
}

#[test]
fn case_enum_two_labels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSuit=(Clubs,Diamonds,Hearts,Spades); var s:TSuit; begin s:=Hearts; case s of Clubs,Diamonds:WriteLn('red0'); Hearts,Spades:WriteLn('black0'); end; end."#
        ),
        &["black0"]
    );
}

#[test]
fn case_enum_comma_list() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay=(Mon,Tue,Wed,Thu,Fri,Sat,Sun); var d:TDay; begin d:=Fri; case d of Mon,Tue,Wed,Thu,Fri:WriteLn('weekday'); Sat,Sun:WriteLn('weekend'); end; end."#
        ),
        &["weekday"]
    );
}

#[test]
fn case_enum_weekend() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay=(Mon,Tue,Wed,Thu,Fri,Sat,Sun); var d:TDay; begin d:=Sun; case d of Mon,Tue,Wed,Thu,Fri:WriteLn('weekday'); Sat,Sun:WriteLn('weekend'); end; end."#
        ),
        &["weekend"]
    );
}

#[test]
fn case_enum_ordinal_based() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLevel=(Low,Med,High); var l:TLevel; begin l:=Med; case l of Low:WriteLn('1'); Med:WriteLn('2'); High:WriteLn('3'); end; end."#
        ),
        &["2"]
    );
}

#[test]
fn case_char_in_subrange_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDigit='0'..'9'; var d:TDigit; begin d:='4'; case d of '0'..'4':WriteLn('low'); '5'..'9':WriteLn('high'); end; end."#
        ),
        &["low"]
    );
}

#[test]
fn case_char_subrange_high() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDigit='0'..'9'; var d:TDigit; begin d:='8'; case d of '0'..'4':WriteLn('low'); '5'..'9':WriteLn('high'); end; end."#
        ),
        &["high"]
    );
}

#[test]
fn case_enum_nested_in_if() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMode=(Read,Write,Append); var m:TMode; begin m:=Write; if m=Write then case m of Read:WriteLn('r'); Write:WriteLn('w'); Append:WriteLn('a'); end; end."#
        ),
        &["w"]
    );
}

#[test]
fn case_char_with_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='q'; case c of 'a'..'d':WriteLn('ad'); 'e'..'h':WriteLn('eh'); else WriteLn('rest'); end; end."#
        ),
        &["rest"]
    );
}

#[test]
fn case_char_mixed_list_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='b'; case c of 'a','c'..'e':WriteLn('hit'); else WriteLn('miss'); end; end."#
        ),
        &["hit"]
    );
}

#[test]
fn case_enum_three_way() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSize=(Small,Medium,Large); var s:TSize; begin s:=Large; case s of Small:WriteLn('s'); Medium:WriteLn('m'); Large:WriteLn('l'); end; end."#
        ),
        &["l"]
    );
}

#[test]
fn case_char_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='0'; case c of '0':WriteLn('zero'); '1'..'9':WriteLn('nz'); end; end."#
        ),
        &["zero"]
    );
}

#[test]
fn case_char_newline() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:=#10; case c of #10:WriteLn('nl'); #13:WriteLn('cr'); else WriteLn('x'); end; end."#
        ),
        &["nl"]
    );
}

#[test]
fn case_enum_first_ordinal() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(Alpha,Beta,Gamma); var x:T; begin x:=Alpha; case x of Alpha:WriteLn(Ord(x)); Beta:WriteLn(1); Gamma:WriteLn(2); end; end."#
        ),
        &["0"]
    );
}

#[test]
fn case_enum_last_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(Alpha,Beta,Gamma); var x:T; begin x:=Gamma; case x of Alpha:WriteLn(0); Beta:WriteLn(1); Gamma:WriteLn(Ord(x)); end; end."#
        ),
        &["2"]
    );
}

#[test]
fn case_char_hex_escape() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:=#65; case c of 'A'..'Z':WriteLn('AZ'); else WriteLn('no'); end; end."#
        ),
        &["AZ"]
    );
}

#[test]
fn case_char_plus_sign() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='+'; case c of '+','-':WriteLn('sign'); '*','/':WriteLn('mul'); else WriteLn('?'); end; end."#
        ),
        &["sign"]
    );
}

#[test]
fn case_enum_scoped_style() {
    assert_eq!(
        run_pascal(
            r#"program T; type TStatus=(Idle,Running,Done); var s:TStatus; begin s:=Running; case s of Idle:WriteLn('i'); Running:WriteLn('run'); Done:WriteLn('d'); end; end."#
        ),
        &["run"]
    );
}

#[test]
fn case_char_range_boundary() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='m'; case c of 'a'..'m':WriteLn('am'); 'n'..'z':WriteLn('nz'); end; end."#
        ),
        &["am"]
    );
}

#[test]
fn case_char_boundary_n() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='n'; case c of 'a'..'m':WriteLn('am'); 'n'..'z':WriteLn('nz'); end; end."#
        ),
        &["nz"]
    );
}

#[test]
fn case_enum_pair_group_a() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(A,B,C,D); var v:T; begin v:=B; case v of A,B:WriteLn('ab'); C,D:WriteLn('cd'); end; end."#
        ),
        &["ab"]
    );
}

#[test]
fn case_enum_pair_group_cd() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(A,B,C,D); var v:T; begin v:=D; case v of A,B:WriteLn('ab'); C,D:WriteLn('cd'); end; end."#
        ),
        &["cd"]
    );
}

#[test]
fn case_char_star_wildcard() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='*'; case c of '*','?':WriteLn('wild'); else WriteLn('lit'); end; end."#
        ),
        &["wild"]
    );
}

#[test]
fn case_enum_default_via_else() {
    assert_eq!(
        run_pascal(
            r#"program T; type TOp=(Add,Sub,Mul,Div); var o:TOp; begin o:=Div; case o of Add:WriteLn('+'); Sub:WriteLn('-'); else WriteLn('other'); end; end."#
        ),
        &["other"]
    );
}

#[test]
fn case_char_comma_three_ranges() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='2'; case c of '0'..'2','4'..'6','8'..'9':WriteLn('grp'); else WriteLn('gap'); end; end."#
        ),
        &["grp"]
    );
}

#[test]
fn case_char_gap_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='3'; case c of '0'..'2','4'..'6','8'..'9':WriteLn('grp'); else WriteLn('gap'); end; end."#
        ),
        &["gap"]
    );
}

#[test]
fn case_enum_in_for_dispatch() {
    assert_eq!(
        run_pascal(
            r#"program T; type T=(On,Off); var s:T; begin s:=On; case s of On:WriteLn('1'); Off:WriteLn('0'); end; s:=Off; case s of On:WriteLn('1'); Off:WriteLn('0'); end; end."#
        ),
        &["1", "0"]
    );
}
