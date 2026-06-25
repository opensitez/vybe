//! `unsafe` blocks, fixed pointers, and stackalloc.
use super::helpers::run_csharp;

#[test]
fn unsafe_block_reads_value_via_pointer() {
    assert_eq!(
        run_csharp(r#"unsafe{
    int x=42;
    int* p=&x;
    Console.WriteLine(*p);
}"#),
        &["42"]
    );
}

#[test]
fn unsafe_pointer_write_mutates_original_variable() {
    assert_eq!(
        run_csharp(r#"unsafe{
    int x=1;
    int* p=&x;
    *p=99;
    Console.WriteLine(x);
}"#),
        &["99"]
    );
}

#[test]
fn fixed_statement_pins_array_for_pointer_arithmetic() {
    assert_eq!(
        run_csharp(r#"int[] arr={10,20,30};
unsafe{
    fixed(int* p=arr){
        Console.WriteLine(*(p+1));
    }
}"#),
        &["20"]
    );
}

#[test]
fn stackalloc_allocates_on_stack_and_is_readable() {
    assert_eq!(
        run_csharp(r#"unsafe{
    int* buf=stackalloc int[3]{1,2,3};
    Console.WriteLine(buf[2]);
}"#),
        &["3"]
    );
}
