//! Actual parsed PHP array_slice Source. Expected outputs independently captured
//! from installed native PHP; complete key/value digests also prove no input mutation.
use vybe_runtime::{Chunk, VM, Value};
pub fn compile_source(source: &str) -> Vec<Chunk> {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap()
}
pub fn verify_vm(source: &str, expected: &str) -> Vec<Chunk> {
    let chunks = compile_source(source);
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    assert_eq!(
        vm.run(chunks.clone()).unwrap(),
        Value::String(expected.into()),
        "native PHP oracle; slice and unchanged original"
    );
    chunks
}
pub const CASES: &[(&str, &str, &str)] = &[
    (
        "middle-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,4,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "0,08,1,z:N7,S08,N2,Z|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "middle-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,4,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "7,08,-2,z:N7,S08,N2,Z|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "negative-offset-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-4,3,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "08,0,z:S08,N2,Z|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "negative-offset-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-4,3,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "08,-2,z:S08,N2,Z|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "negative-length-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,-1,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "0,08,1,z:N7,S08,N2,Z|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "negative-length-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,-1,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "7,08,-2,z:N7,S08,N2,Z|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "both-negative-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-5,-2,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "0,08,1:N7,S08,N2|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "both-negative-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-5,-2,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "7,08,-2:N7,S08,N2|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "past-end-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,20,2,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "past-end-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,20,2,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "zero-length-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,0,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "zero-length-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,0,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "very-negative-offset-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-20,2,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "s,0:S,N7|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "very-negative-offset-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-20,2,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "s,7:S,N7|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "very-negative-length-reindex",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,-20,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "very-negative-length-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,1,-20,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "omitted",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,2);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "08,0,z,1:S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "null-length",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-3,null);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "0,z,1:N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "null-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['s'=>'S',7=>'N7','08'=>'S08',-2=>'N2','z'=>'Z',4=>'N4'];$before=digest_slice_input($input);$slice=array_slice($input,-3,null,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "-2,z,4:N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4|s,7,08,-2,z,4:S,N7,S08,N2,Z,N4",
    ),
    (
        "dense",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['A','B','C','D'];$before=digest_slice_input($input);$slice=array_slice($input,-3,2,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "0,1:B,C|0,1,2,3:A,B,C,D|0,1,2,3:A,B,C,D",
    ),
    (
        "dense-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['A','B','C','D'];$before=digest_slice_input($input);$slice=array_slice($input,1,2,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "1,2:B,C|0,1,2,3:A,B,C,D|0,1,2,3:A,B,C,D",
    ),
    (
        "numeric-string",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['8'=>'N8','08'=>'S08','k'=>'K',12=>'N12'];$before=digest_slice_input($input);$slice=array_slice($input,0,4,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "0,08,k,1:N8,S08,K,N12|8,08,k,12:N8,S08,K,N12|8,08,k,12:N8,S08,K,N12",
    ),
    (
        "numeric-string-keep",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=['8'=>'N8','08'=>'S08','k'=>'K',12=>'N12'];$before=digest_slice_input($input);$slice=array_slice($input,0,4,true);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        "8,08,k,12:N8,S08,K,N12|8,08,k,12:N8,S08,K,N12|8,08,k,12:N8,S08,K,N12",
    ),
    (
        "empty",
        "<?php function digest_slice_input($a){return implode(',',array_keys($a)).':'.implode(',',array_values($a));}$input=[];$before=digest_slice_input($input);$slice=array_slice($input,0,4,false);return digest_slice_input($slice).'|'.$before.'|'.digest_slice_input($input);",
        ":|:|:",
    ),
];
#[test]
fn actual_source_middle_reindex() {
    let (_, source, expected) = CASES[0];
    verify_vm(source, expected);
}
#[test]
fn actual_source_middle_keep() {
    let (_, source, expected) = CASES[1];
    verify_vm(source, expected);
}
#[test]
fn actual_source_negative_offset_reindex() {
    let (_, source, expected) = CASES[2];
    verify_vm(source, expected);
}
#[test]
fn actual_source_negative_offset_keep() {
    let (_, source, expected) = CASES[3];
    verify_vm(source, expected);
}
#[test]
fn actual_source_negative_length_reindex() {
    let (_, source, expected) = CASES[4];
    verify_vm(source, expected);
}
#[test]
fn actual_source_negative_length_keep() {
    let (_, source, expected) = CASES[5];
    verify_vm(source, expected);
}
#[test]
fn actual_source_both_negative_reindex() {
    let (_, source, expected) = CASES[6];
    verify_vm(source, expected);
}
#[test]
fn actual_source_both_negative_keep() {
    let (_, source, expected) = CASES[7];
    verify_vm(source, expected);
}
#[test]
fn actual_source_past_end_reindex() {
    let (_, source, expected) = CASES[8];
    verify_vm(source, expected);
}
#[test]
fn actual_source_past_end_keep() {
    let (_, source, expected) = CASES[9];
    verify_vm(source, expected);
}
#[test]
fn actual_source_zero_length_reindex() {
    let (_, source, expected) = CASES[10];
    verify_vm(source, expected);
}
#[test]
fn actual_source_zero_length_keep() {
    let (_, source, expected) = CASES[11];
    verify_vm(source, expected);
}
#[test]
fn actual_source_very_negative_offset_reindex() {
    let (_, source, expected) = CASES[12];
    verify_vm(source, expected);
}
#[test]
fn actual_source_very_negative_offset_keep() {
    let (_, source, expected) = CASES[13];
    verify_vm(source, expected);
}
#[test]
fn actual_source_very_negative_length_reindex() {
    let (_, source, expected) = CASES[14];
    verify_vm(source, expected);
}
#[test]
fn actual_source_very_negative_length_keep() {
    let (_, source, expected) = CASES[15];
    verify_vm(source, expected);
}
#[test]
fn actual_source_omitted() {
    let (_, source, expected) = CASES[16];
    verify_vm(source, expected);
}
#[test]
fn actual_source_null_length() {
    let (_, source, expected) = CASES[17];
    verify_vm(source, expected);
}
#[test]
fn actual_source_null_keep() {
    let (_, source, expected) = CASES[18];
    verify_vm(source, expected);
}
#[test]
fn actual_source_dense() {
    let (_, source, expected) = CASES[19];
    verify_vm(source, expected);
}
#[test]
fn actual_source_dense_keep() {
    let (_, source, expected) = CASES[20];
    verify_vm(source, expected);
}
#[test]
fn actual_source_numeric_string() {
    let (_, source, expected) = CASES[21];
    verify_vm(source, expected);
}
#[test]
fn actual_source_numeric_string_keep() {
    let (_, source, expected) = CASES[22];
    verify_vm(source, expected);
}
#[test]
fn actual_source_empty() {
    let (_, source, expected) = CASES[23];
    verify_vm(source, expected);
}
