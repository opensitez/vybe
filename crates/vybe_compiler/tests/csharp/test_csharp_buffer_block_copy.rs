//! `Buffer.BlockCopy` copies a byte range between primitive arrays.
use super::helpers::run_csharp;

#[test]
fn buffer_block_copy_transfers_bytes_between_int_arrays() {
    assert_eq!(
        run_csharp(
            r#"
int[] source = { 0x01020304, 0 };
int[] dest = { 0, 0 };
System.Buffer.BlockCopy(source, 0, dest, 0, 4);
Console.WriteLine(dest[0]);
"#
        ),
        &["67305985"]
    );
}
