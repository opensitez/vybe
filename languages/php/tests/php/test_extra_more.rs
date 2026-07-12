use super::helpers::run_prints;

// ── String ────────────────────────────────────────────────────

#[test]
fn nl2br_crlf() {
    assert_eq!(
        run_prints("<?php echo nl2br(\"a\\r\\nb\"); "),
        vec!["a<br />\r\nb"]
    );
}
#[test]
fn str_rot13_encode_decode() {
    assert_eq!(
        run_prints(r#"<?php $s='Hello World'; echo str_rot13(str_rot13($s))===$s?'ok':'fail'; "#),
        vec!["ok"]
    );
}
#[test]
fn quoted_printable_encode() {
    assert_eq!(
        run_prints(
            r#"<?php $e=quoted_printable_encode('Subject: =?UTF-8'); echo strlen($e)>0?'ok':'fail'; "#
        ),
        vec!["ok"]
    );
}
#[test]
fn addslashes_removes_with_stripslashes() {
    assert_eq!(
        run_prints(r#"<?php echo stripslashes(addslashes("it's a \"test\"")); "#),
        vec!["it's a \"test\""]
    );
}
#[test]
fn wordwrap_no_cut() {
    assert_eq!(
        run_prints(r#"<?php echo wordwrap('The quick',5,"\n",false); "#),
        vec!["The", "quick"]
    );
}
#[test]
fn chunk_split_empty_string() {
    assert_eq!(
        run_prints(r#"<?php echo chunk_split('',2,':'); "#),
        vec![":"]
    );
}
#[test]
fn hex2bin_bin2hex() {
    assert_eq!(
        run_prints(r#"<?php echo hex2bin(bin2hex('AB')); "#),
        vec!["AB"]
    );
}
#[test]
fn count_chars_mode0() {
    assert_eq!(
        run_prints(r#"<?php $c=count_chars('hello',1); echo $c[ord('l')]; "#),
        vec!["2"]
    );
}
#[test]
fn substr_replace_array() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',',substr_replace(['a','bb','ccc'],'-',1)); "#),
        vec!["a-,b-,c-"]
    );
}
#[test]
fn str_getcsv_basic() {
    assert_eq!(
        run_prints(r#"<?php echo implode('|',str_getcsv('a,b,c')); "#),
        vec!["a|b|c"]
    );
}

// ── Array ─────────────────────────────────────────────────────

#[test]
fn array_map_keys_preserved_assoc() {
    assert_eq!(
        run_prints(
            r#"<?php $r=array_map(fn($v)=>$v*2,['x'=>3,'y'=>5]); echo $r['x'].','.$r['y']; "#
        ),
        vec!["6,10"]
    );
}
#[test]
fn array_reduce_empty_returns_initial() {
    assert_eq!(
        run_prints(r#"<?php echo array_reduce([],fn($c,$v)=>$c+$v,99); "#),
        vec!["99"]
    );
}
#[test]
fn array_filter_default_removes_all_falsy() {
    assert_eq!(
        run_prints(r#"<?php echo count(array_filter([0,'',null,false,[]])); "#),
        vec!["0"]
    );
}
#[test]
fn array_intersect_key() {
    assert_eq!(
        run_prints(
            r#"<?php $r=array_intersect_key(['a'=>1,'b'=>2,'c'=>3],['a'=>99,'c'=>99]); echo implode(',',array_keys($r)); "#
        ),
        vec!["a,c"]
    );
}
#[test]
fn array_combine_unequal_throws() {
    assert_eq!(
        run_prints(r#"<?php try{array_combine([1,2],[1]);}catch(ValueError $e){echo 'err';} "#),
        vec!["err"]
    );
}
#[test]
fn array_multisort_desc() {
    assert_eq!(
        run_prints(r#"<?php $a=[3,1,2]; array_multisort($a,SORT_DESC); echo implode(',',$a); "#),
        vec!["3,2,1"]
    );
}
#[test]
fn array_unique_type_string() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',',array_unique([1,'1',1.0],SORT_STRING)); "#),
        vec!["1"]
    );
}
#[test]
fn array_key_exists_null_value() {
    assert_eq!(
        run_prints(r#"<?php $a=['k'=>null]; echo array_key_exists('k',$a)?'yes':'no'; "#),
        vec!["yes"]
    );
}

// ── OOP ───────────────────────────────────────────────────────

#[test]
fn interface_extends_multiple() {
    assert_eq!(
        run_prints(
            r#"<?php
interface X2{public function x():int;}
interface Y2{public function y():int;}
interface Z2 extends X2,Y2{}
class Impl implements Z2{public function x():int{return 1;}public function y():int{return 2;}}
$o=new Impl; echo $o->x()+$o->y();
"#
        ),
        vec!["3"]
    );
}
#[test]
fn trait_overrides_parent_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class P{public function hello():string{return 'parent';}}
trait T{public function hello():string{return 'trait';}}
class C extends P{use T;}
echo (new C)->hello();
"#
        ),
        vec!["trait"]
    );
}
#[test]
fn anonymous_class_instanceof_parent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base{}
$o=new class extends Base{};
echo $o instanceof Base?'yes':'no';
"#
        ),
        vec!["yes"]
    );
}
#[test]
fn object_spread_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class A{public int $x=1;public int $y=2;}
$o=new A;
$arr=(array)$o;
echo $arr['x'].','.$arr['y'];
"#
        ),
        vec!["1,2"]
    );
}
#[test]
fn clone_triggers_clone_magic() {
    assert_eq!(
        run_prints(
            r#"<?php
class C{public int $v=0;public function __clone(){$this->v++;}}
$a=new C; $b=clone $a;
echo $b->v;
"#
        ),
        vec!["1"]
    );
}

// ── Math / numbers ────────────────────────────────────────────

#[test]
fn random_int_single_value_range() {
    assert_eq!(run_prints(r#"<?php echo random_int(42,42); "#), vec!["42"]);
}
#[test]
fn base_convert_binary_to_hex() {
    assert_eq!(
        run_prints(r#"<?php echo base_convert('11111111',2,16); "#),
        vec!["ff"]
    );
}
#[test]
fn math_constants_relation() {
    assert_eq!(
        run_prints(r#"<?php echo round(M_LN10/M_LN2,4); "#),
        vec!["3.3219"]
    );
}
#[test]
fn intdiv_negative_result() {
    assert_eq!(run_prints(r#"<?php echo intdiv(-9,2); "#), vec!["-4"]);
}
#[test]
fn fmod_preserves_sign() {
    assert_eq!(run_prints(r#"<?php echo fmod(10.5,-3.0); "#), vec!["1.5"]);
}

// ── Type juggling ─────────────────────────────────────────────

#[test]
fn int_overflow_to_float() {
    assert_eq!(
        run_prints(r#"<?php $v=PHP_INT_MAX+1; echo gettype($v); "#),
        vec!["double"]
    );
}
#[test]
fn string_concatenation_with_null() {
    assert_eq!(
        run_prints(r#"<?php $s='hello'; $s.=null; echo $s; "#),
        vec!["hello"]
    );
}
#[test]
fn empty_array_vs_null() {
    assert_eq!(
        run_prints(r#"<?php echo ([]===null)?'eq':'neq'; "#),
        vec!["neq"]
    );
}
#[test]
fn strict_false_not_null() {
    assert_eq!(
        run_prints(r#"<?php echo (false===null)?'eq':'neq'; "#),
        vec!["neq"]
    );
}
#[test]
fn intval_octal_string() {
    assert_eq!(run_prints(r#"<?php echo intval('0777',8); "#), vec!["511"]);
}

// ── Control flow ──────────────────────────────────────────────

#[test]
fn match_returns_value() {
    assert_eq!(
        run_prints(r#"<?php $x=match(2){1=>'one',2=>'two',3=>'three'}; echo $x; "#),
        vec!["two"]
    );
}
#[test]
fn foreach_nested_break2() {
    assert_eq!(
        run_prints(
            r#"<?php
$found='';
foreach(['a','b'] as $i) {
    foreach([1,2,3] as $j) {
        if($j===2){$found=$i.$j;break 2;}
    }
}
echo $found;
"#
        ),
        vec!["a2"]
    );
}
#[test]
fn do_while_condition_last() {
    assert_eq!(
        run_prints(
            r#"<?php
$x=5; $count=0;
do{$x--;$count++;}while($x>0);
echo $count;
"#
        ),
        vec!["5"]
    );
}
#[test]
fn continue_skips_echo() {
    assert_eq!(
        run_prints(r#"<?php for($i=1;$i<=5;$i++){if($i===3)continue;echo $i;} "#),
        vec!["1245"]
    );
}

// ── Misc ─────────────────────────────────────────────────────

#[test]
fn array_key_first_empty_null() {
    assert_eq!(
        run_prints(r#"<?php var_export(array_key_first([])); "#),
        vec!["NULL"]
    );
}
#[test]
fn array_key_last_empty_null() {
    assert_eq!(
        run_prints(r#"<?php var_export(array_key_last([])); "#),
        vec!["NULL"]
    );
}

// ── Exceptions ────────────────────────────────────────────────

#[test]
fn catch_multiple_types_php8() {
    assert_eq!(
        run_prints(
            r#"<?php
try{throw new InvalidArgumentException('iae');}
catch(RuntimeException|InvalidArgumentException $e){echo $e->getMessage();}
"#
        ),
        vec!["iae"]
    );
}
#[test]
fn exception_code_zero_default() {
    assert_eq!(
        run_prints(
            r#"<?php try{throw new Exception('e');}catch(Exception $e){echo $e->getCode();} "#
        ),
        vec!["0"]
    );
}
#[test]
fn rethrow_preserves_original() {
    assert_eq!(
        run_prints(
            r#"<?php
try{
    try{throw new RuntimeException('orig',404);}
    catch(RuntimeException $e){throw new LogicException('wrap',0,$e);}
}catch(LogicException $e){echo $e->getPrevious()->getCode();}
"#
        ),
        vec!["404"]
    );
}
