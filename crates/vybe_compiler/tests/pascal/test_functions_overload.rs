/// Overload resolution and default parameter combinations.
use super::helpers::run_pascal;

#[test]
fn overload_three_int_adders() {
    assert_eq!(
        run_pascal(r#"program T; function F(n:Integer):Integer; overload; begin Result:=n+1; end; function F(n:Integer;m:Integer):Integer; overload; begin Result:=n+m; end; function F(n,m,k:Integer):Integer; overload; begin Result:=n+m+k; end; begin WriteLn(F(5)); WriteLn(F(2,3)); WriteLn(F(1,2,3)); end."#),
        &["6", "5", "6"]
    );
}

#[test]
fn overload_string_and_integer_write() {
    assert_eq!(
        run_pascal(r#"program T; procedure P(s:string); overload; begin WriteLn(s); end; procedure P(n:Integer); overload; begin WriteLn(n); end; begin P('hi'); P(42); end."#),
        &["hi", "42"]
    );
}

#[test]
fn overload_real_vs_integer_double() {
    assert_eq!(
        run_pascal(r#"program T; function Twice(v:Integer):Integer; overload; begin Result:=v*2; end; function Twice(v:Real):Real; overload; begin Result:=v*2.0; end; begin WriteLn(Twice(3)); WriteLn(Trunc(Twice(2.5))); end."#),
        &["6", "5"]
    );
}

#[test]
fn overload_boolean_inverter_variants() {
    assert_eq!(
        run_pascal(r#"program T; function Flip(b:Boolean):Boolean; overload; begin Result:=not b; end; function Flip(n:Integer):Integer; overload; begin Result:=-n; end; begin WriteLn(Flip(true)); WriteLn(Flip(7)); end."#),
        &["false", "-7"]
    );
}

#[test]
fn overload_char_vs_string_first() {
    assert_eq!(
        run_pascal(r#"program T; function Head(c:Char):string; overload; begin Result:=c; end; function Head(s:string):string; overload; begin Result:=Copy(s,1,1); end; begin WriteLn(Head('z')); WriteLn(Head('abc')); end."#),
        &["z", "a"]
    );
}

#[test]
fn overload_procedure_zero_and_one_arg() {
    assert_eq!(
        run_pascal(r#"program T; procedure Ping; overload; begin WriteLn('p'); end; procedure Ping(n:Integer); overload; begin WriteLn(n); end; begin Ping; Ping(9); end."#),
        &["p", "9"]
    );
}

#[test]
fn overload_array_fixed_vs_open() {
    assert_eq!(
        run_pascal(r#"program T; function Sum(const a:array of Integer):Integer; overload; var i:Integer; begin Result:=0; for i:=0 to High(a) do Result:=Result+a[i]; end; function Sum(a:array[1..2] of Integer):Integer; begin Result:=a[1]+a[2]; end; var b:array[1..2] of Integer; begin b[1]:=4; b[2]:=6; WriteLn(Sum(b)); end."#),
        &["10"]
    );
}

#[test]
fn overload_resolution_prefers_exact_count() {
    assert_eq!(
        run_pascal(r#"program T; function G(x:Integer):Integer; overload; begin Result:=x; end; function G(x,y:Integer):Integer; overload; begin Result:=x+y; end; begin WriteLn(G(4)); WriteLn(G(2,5)); end."#),
        &["4", "7"]
    );
}

#[test]
fn default_single_int_param_added() {
    assert_eq!(
        run_pascal(r#"program T; function Add(n:Integer; d:Integer=3):Integer; begin Result:=n+d; end; begin WriteLn(Add(10)); WriteLn(Add(10,7)); end."#),
        &["13", "17"]
    );
}

#[test]
fn default_string_prefix_empty() {
    assert_eq!(
        run_pascal(r#"program T; function Tag(s:string; pfx:string='>'):string; begin Result:=pfx+s; end; begin WriteLn(Tag('x')); WriteLn(Tag('y','!')); end."#),
        &[">x", "!y"]
    );
}

#[test]
fn default_two_params_second_only() {
    assert_eq!(
        run_pascal(r#"program T; function Mul(a:Integer; b:Integer=2; c:Integer=10):Integer; begin Result:=a*b+c; end; begin WriteLn(Mul(5)); WriteLn(Mul(5,3)); end."#),
        &["20", "25"]
    );
}

#[test]
fn default_bool_verbose_flag() {
    assert_eq!(
        run_pascal(r#"program T; procedure Show(n:Integer; loud:Boolean=false); begin if loud then WriteLn('L'+IntToStr(n)) else WriteLn(IntToStr(n)); end; begin Show(1); Show(2,true); end."#),
        &["1", "L2"]
    );
}

#[test]
fn default_real_tolerance_compare() {
    assert_eq!(
        run_pascal(r#"program T; function Near(a,b:Real; eps:Real=0.01):Boolean; begin Result:=Abs(a-b)<=eps; end; begin WriteLn(Near(1.0,1.005)); WriteLn(Near(1.0,1.1,0.05)); end."#),
        &["true", "false"]
    );
}

#[test]
fn default_char_pad_fill() {
    assert_eq!(
        run_pascal(r#"program T; function Pad(c:Char; n:Integer=3):string; var i:Integer; begin Result:=''; for i:=1 to n do Result:=Result+c; end; begin WriteLn(Pad('-')); WriteLn(Pad('*',2)); end."#),
        &["---", "**"]
    );
}

#[test]
fn default_nested_call_uses_inner_default() {
    assert_eq!(
        run_pascal(r#"program T; function IncBy(n:Integer; d:Integer=5):Integer; begin Result:=n+d; end; begin WriteLn(IncBy(IncBy(1))); end."#),
        &["11"]
    );
}

#[test]
fn default_explicit_zero_overrides() {
    assert_eq!(
        run_pascal(r#"program T; function Scale(n:Integer; k:Integer=2):Integer; begin Result:=n*k; end; begin WriteLn(Scale(7,0)); end."#),
        &["0"]
    );
}

#[test]
fn default_mixed_types_int_and_string() {
    assert_eq!(
        run_pascal(r#"program T; function LabelOf(n:Integer; lbl:string='n'):string; begin Result:=lbl+IntToStr(n); end; begin WriteLn(LabelOf(4)); WriteLn(LabelOf(4,'v')); end."#),
        &["n4", "v4"]
    );
}

#[test]
fn overload_plus_default_on_one_variant() {
    assert_eq!(
        run_pascal(r#"program T; function Val(n:Integer; bias:Integer=0):Integer; overload; begin Result:=n+bias; end; function Val(s:string):Integer; overload; begin Result:=Length(s); end; begin WriteLn(Val(3)); WriteLn(Val('abc')); end."#),
        &["3", "3"]
    );
}

#[test]
fn default_param_expression_in_call() {
    assert_eq!(
        run_pascal(r#"program T; function Offset(n:Integer; d:Integer=1):Integer; begin Result:=n+d; end; var x:Integer; begin x:=4; WriteLn(Offset(x,x)); end."#),
        &["8"]
    );
}

#[test]
fn overload_three_string_concat_styles() {
    assert_eq!(
        run_pascal(r#"program T; function Join(a,b:string):string; overload; begin Result:=a+b; end; function Join(a,b,c:string):string; overload; begin Result:=a+b+c; end; begin WriteLn(Join('a','b')); WriteLn(Join('x','y','z')); end."#),
        &["ab", "xyz"]
    );
}

#[test]
fn default_all_three_middle_specified() {
    assert_eq!(
        run_pascal(r#"program T; function F(a,b:Integer=1; c:Integer=2):Integer; begin Result:=a+b+c; end; begin WriteLn(F(10,20)); end."#),
        &["32"]
    );
}

#[test]
fn overload_integer_set_membership() {
    assert_eq!(
        run_pascal(r#"program T; type TD=(A,B); function Has(v:TD):Boolean; overload; begin Result:=v=B; end; function Has(n:Integer):Boolean; overload; begin Result:=n>0; end; begin WriteLn(Has(B)); WriteLn(Has(-1)); end."#),
        &["true", "false"]
    );
}

#[test]
fn default_procedure_optional_message() {
    assert_eq!(
        run_pascal(r#"program T; procedure Log(msg:string='done'); begin WriteLn(msg); end; begin Log; Log('go'); end."#),
        &["done", "go"]
    );
}

#[test]
fn overload_record_and_integer() {
    assert_eq!(
        run_pascal(r#"program T; type TPt=record X,Y:Integer; end; function Area(p:TPt):Integer; overload; begin Result:=p.X*p.Y; end; function Area(w,h:Integer):Integer; overload; begin Result:=w*h; end; var p:TPt; begin p.X:=3; p.Y:=4; WriteLn(Area(p)); WriteLn(Area(2,5)); end."#),
        &["12", "10"]
    );
}

#[test]
fn default_negative_literal() {
    assert_eq!(
        run_pascal(r#"program T; function Shift(n:Integer; d:Integer=-1):Integer; begin Result:=n+d; end; begin WriteLn(Shift(5)); end."#),
        &["4"]
    );
}

#[test]
fn overload_enum_ord_helper() {
    assert_eq!(
        run_pascal(r#"program T; type TS=(One,Two,Three); function N(v:TS):Integer; overload; begin Result:=Ord(v); end; function N(i:Integer):Integer; overload; begin Result:=i*10; end; begin WriteLn(N(Two)); WriteLn(N(2)); end."#),
        &["1", "20"]
    );
}

#[test]
fn default_chain_three_levels() {
    assert_eq!(
        run_pascal(r#"program T; function A(n:Integer=1):Integer; begin Result:=n; end; function B(n:Integer=2):Integer; begin Result:=n; end; begin WriteLn(A(B)); end."#),
        &["2"]
    );
}

#[test]
fn overload_pointer_nil_check() {
    assert_eq!(
        run_pascal(r#"program T; function Ok(p:Pointer):Boolean; overload; begin Result:=p=nil; end; function Ok(n:Integer):Boolean; overload; begin Result:=n=0; end; begin WriteLn(Ok(nil)); WriteLn(Ok(1)); end."#),
        &["true", "false"]
    );
}

#[test]
fn default_string_with_concat_call() {
    assert_eq!(
        run_pascal(r#"program T; function Wrap(s:string; open:string='['; close:string=']'):string; begin Result:=open+s+close; end; begin WriteLn(Wrap('a')); WriteLn(Wrap('b','(',' )')); end."#),
        &["[a]", "(b)"]
    );
}

#[test]
fn overload_two_procedures_by_name_count() {
    assert_eq!(
        run_pascal(r#"program T; procedure Emit; overload; begin WriteLn('0'); end; procedure Emit(c:Char); overload; begin WriteLn(c); end; procedure Emit(s:string); overload; begin WriteLn(s); end; begin Emit; Emit('x'); Emit('ok'); end."#),
        &["0", "x", "ok"]
    );
}

#[test]
fn default_integer_max_bound() {
    assert_eq!(
        run_pascal(r#"program T; function Cap(n:Integer; m:Integer=100):Integer; begin if n>m then Result:=m else Result:=n; end; begin WriteLn(Cap(150)); WriteLn(Cap(80,50)); end."#),
        &["100", "50"]
    );
}

#[test]
fn overload_function_result_types_int_string() {
    assert_eq!(
        run_pascal(r#"program T; function Desc(n:Integer):Integer; overload; begin Result:=n; end; function Desc(s:string):string; overload; begin Result:=s; end; begin WriteLn(Desc(7)); WriteLn(Desc('z')); end."#),
        &["7", "z"]
    );
}

#[test]
fn default_bool_true_literal() {
    assert_eq!(
        run_pascal(r#"program T; function Flag(on:Boolean=true):string; begin if on then Result:='y' else Result:='n'; end; begin WriteLn(Flag); WriteLn(Flag(false)); end."#),
        &["y", "n"]
    );
}

#[test]
fn overload_var_param_vs_value() {
    assert_eq!(
        run_pascal(r#"program T; procedure Bump(var n:Integer); overload; begin Inc(n); end; procedure Bump(n:Integer); overload; begin WriteLn(n+1); end; var x:Integer; begin x:=1; Bump(x); WriteLn(x); Bump(4); end."#),
        &["2", "5"]
    );
}

#[test]
fn default_real_multiplier() {
    assert_eq!(
        run_pascal(r#"program T; function ScaleR(v:Real; k:Real=2.5):Real; begin Result:=v*k; end; begin WriteLn(Trunc(ScaleR(4.0))); end."#),
        &["10"]
    );
}

#[test]
fn overload_case_insensitive_tag() {
    assert_eq!(
        run_pascal(r#"program T; function Kind(c:Char):Integer; overload; begin Result:=Ord(c); end; function Kind(s:string):Integer; overload; begin Result:=Length(s); end; begin WriteLn(Kind('A')); WriteLn(Kind('ab')); end."#),
        &["65", "2"]
    );
}

#[test]
fn default_two_strings_second_default() {
    assert_eq!(
        run_pascal(r#"program T; function Pair(a:string; b:string='b'):string; begin Result:=a+b; end; begin WriteLn(Pair('a')); WriteLn(Pair('x','y')); end."#),
        &["ab", "xy"]
    );
}

#[test]
fn overload_nested_integer_calls() {
    assert_eq!(
        run_pascal(r#"program T; function Id(n:Integer):Integer; overload; begin Result:=n; end; function Id(a,b:Integer):Integer; overload; begin Result:=a+b; end; begin WriteLn(Id(Id(2),Id(3))); end."#),
        &["5"]
    );
}

#[test]
fn default_with_arithmetic_body() {
    assert_eq!(
        run_pascal(r#"program T; function Pow2(n:Integer; e:Integer=2):Integer; var i:Integer; begin Result:=1; for i:=1 to e do Result:=Result*n; end; begin WriteLn(Pow2(3)); WriteLn(Pow2(2,3)); end."#),
        &["9", "8"]
    );
}

#[test]
fn overload_const_param_string() {
    assert_eq!(
        run_pascal(r#"program T; function Len(const s:string):Integer; overload; begin Result:=Length(s); end; function Len(c:Char):Integer; overload; begin Result:=1; end; begin WriteLn(Len('abc')); WriteLn(Len('z')); end."#),
        &["3", "1"]
    );
}

#[test]
fn default_triple_int_only_first_passed() {
    assert_eq!(
        run_pascal(r#"program T; function Sum3(a:Integer; b:Integer=1; c:Integer=2):Integer; begin Result:=a+b+c; end; begin WriteLn(Sum3(10)); end."#),
        &["13"]
    );
}

#[test]
fn overload_and_default_combined() {
    assert_eq!(
        run_pascal(r#"program T; function F(n:Integer):Integer; overload; begin Result:=n; end; function F(n:Integer; k:Integer=2):Integer; overload; begin Result:=n*k; end; begin WriteLn(F(5)); WriteLn(F(5,3)); end."#),
        &["5", "15"]
    );
}
