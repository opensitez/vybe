//! Binary-format conformance for the **function-references** proposal.
//! Spec: `proposals/function-references/proposals/function-references/Overview.md`
//!
//! Normative opcodes (Overview.md "Binary format" table):
//!
//! | 0x14 | `call_ref $t`        | `$t : u32` |
//! | 0x15 | `return_call_ref $t` | `$t : u32` |
//! | 0xd4 | `ref.as_non_null`    |            |
//! | 0xd5 | `br_on_null $l`      | `$l : u32` |
//! | 0xd6 | `br_on_non_null $l`  | `$l : u32` |
//!
//! These assert the *reader* accepts each encoding. `br_on_null` / `br_on_non_null`
//! were previously emitted by the writer but rejected by the reader, so a module
//! using them could not round-trip.

fn write_leb_u32(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn standard_module_with_sections(sections: &[(u8, Vec<u8>)]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    for (id, payload) in sections {
        bytes.push(*id);
        write_leb_u32(&mut bytes, payload.len() as u32);
        bytes.extend_from_slice(payload);
    }
    bytes
}

fn code_section_for_body(body_ops: &[u8]) -> Vec<u8> {
    let mut body = Vec::new();
    body.push(0x00); // no local declarations
    body.extend_from_slice(body_ops);
    body.push(0x0b); // end

    let mut section = Vec::new();
    write_leb_u32(&mut section, 1); // one function body
    write_leb_u32(&mut section, body.len() as u32);
    section.extend_from_slice(&body);
    section
}

/// A module whose single function `() -> i32` runs `body_ops`.
fn module_with_body(body_ops: &[u8]) -> Vec<u8> {
    standard_module_with_sections(&[
        // type: (func (result i32))
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x7f]),
        // func: one function, type 0
        (3, vec![0x01, 0x00]),
        (10, code_section_for_body(body_ops)),
    ])
}

#[test]
fn reader_accepts_br_on_null_encoding() {
    // `br_on_null $l : [t* (ref null ht)] -> [t* (ref ht)]` — branches without
    // the reference when null, otherwise re-types it non-null and falls through.
    // A void block, so the label type `t*` is empty and the branch needs only
    // the reference itself on the stack.
    let wasm = module_with_body(&[
        0x02, 0x40, // block (void)
        0xd0, 0x70, //   ref.null func
        0xd5, 0x00, //   br_on_null 0
        0x1a, //   drop  (non-null fallthrough leaves the ref)
        0x0b, // end
        0x41, 0x01, // i32.const 1
    ]);
    vybe_platform_wasm::read_wasm(&wasm).expect("br_on_null (0xd5) must decode");
}

#[test]
fn reader_accepts_br_on_non_null_encoding() {
    // `br_on_non_null $l : [t* (ref null ht)] -> [t*]` — branches *with* the
    // reference when non-null, otherwise pops it and falls through. The ref
    // is the LAST of the target label's expected values (function-references
    // Overview: "the branch target label must end with a non-null reference
    // type"), so the label CANNOT be void — the old void-block form here was
    // an invalid module a conforming validator rejects.
    let wasm = module_with_body(&[
        0x02, 0x70, // block (result funcref)
        0xd0, 0x70, //   ref.null func
        0xd6, 0x00, //   br_on_non_null 0
        0xd0, 0x70, //   ref.null func (fallthrough block result)
        0x0b, // end
        0x1a, // drop
        0x41, 0x00, // i32.const 0
    ]);
    vybe_platform_wasm::read_wasm(&wasm).expect("br_on_non_null (0xd6) must decode");
}

#[test]
fn reader_accepts_ref_as_non_null_encoding() {
    let wasm = module_with_body(&[
        0xd0, 0x70, // ref.null func
        0xd4, // ref.as_non_null
        0x1a, // drop
        0x41, 0x00, // i32.const 0
    ]);
    vybe_platform_wasm::read_wasm(&wasm).expect("ref.as_non_null (0xd4) must decode");
}

#[test]
fn reader_accepts_call_ref_encoding() {
    // `call_ref $t` takes a *type* index immediate, not a function index.
    let wasm = module_with_body(&[
        0xd0, 0x70, // ref.null func
        0x14, 0x00, // call_ref type 0  -> () -> i32
    ]);
    vybe_platform_wasm::read_wasm(&wasm).expect("call_ref (0x14) must decode");
}

#[test]
fn reader_accepts_return_call_ref_encoding() {
    let wasm = module_with_body(&[
        0xd0, 0x70, // ref.null func
        0x15, 0x00, // return_call_ref type 0
    ]);
    vybe_platform_wasm::read_wasm(&wasm).expect("return_call_ref (0x15) must decode");
}
