//! Native-only actual fixed synchronous Source declaration constructor.
//! Original Chunk instructions and VM entry remain unchanged.
use crate::encoding::*;
use vybe_runtime::{Chunk, chunk::LocalOrigin};
#[derive(Clone, Debug)]
pub struct Entry {
    pub chunk_index: usize,
    pub user_arity: u8,
    pub control_slot: u16,
    pub constructor_func: u32,
    pub body_func: u32,
    pub body_type: u32,
    pub cont_type: u32,
}
pub fn entries(chunks: &[Chunk]) -> Vec<Entry> {
    let mut entries = Vec::new();
    for (i, c) in chunks.iter().enumerate() {
        let mut marker = None;
        for n in &c.local_names {
            if n.origin != LocalOrigin::Compiler {
                continue;
            }
            let Some(fields) = n
                .name
                .strip_prefix("<native-generator-entry:")
                .and_then(|s| s.strip_suffix('>'))
            else {
                continue;
            };
            let v: Vec<u16> = fields
                .split(':')
                .map(|v| v.parse().expect("generator entry integer"))
                .collect();
            assert_eq!(v.len(), 2);
            assert!(marker.is_none(), "duplicate generator entry marker");
            assert!(c.is_generator && !c.is_async && c.func_sig.is_none());
            assert_eq!(n.slot, v[0]);
            assert!(n.slot < c.local_count);
            assert_eq!(u16::from(c.arity), v[1] + 1);
            assert!(c.result_arity <= 1);
            assert!(
                c.local_names
                    .iter()
                    .any(|x| x.origin == LocalOrigin::Compiler
                        && x.name == "__generator_entry_control"
                        && x.slot == n.slot),
                "actual compiler entry control producer"
            );
            marker = Some((v[0], u8::try_from(v[1]).expect("generator user arity")));
        }
        if let Some((slot, arity)) = marker {
            entries.push(Entry {
                chunk_index: i,
                user_arity: arity,
                control_slot: slot,
                constructor_func: 0,
                body_func: 0,
                body_type: 0,
                cont_type: 0,
            });
        // Unmarked rest/async/lambda/method constructors retain their existing lowering.
        }
    }
    entries
}
pub fn body(
    e: &Entry,
    chunks: &[Chunk],
    ctx: &super::types::WasmTypeContext,
    rt: &std::collections::HashMap<(&str, &str), usize>,
) -> Vec<u8> {
    use super::callable::{external, gc, local, typed};
    let mut b = Vec::new();
    // Two genuine locals: owner, then typed continuation. Hidden environment and
    // source bound args are actual constructor parameters and execute no body.
    write_leb128_u32(&mut b, 2);
    write_leb128_u32(&mut b, 1);
    b.push(TYPE_EXTERNREF);
    write_leb128_u32(&mut b, 1);
    b.push(0x63);
    write_leb128_i64(&mut b, ctx.continuation_type_idx as i64);
    let owner = u32::from(e.user_arity) + 1;
    let cont = owner + 1;
    for i in 0..=e.user_arity {
        local(&mut b, 0x20, u32::from(i));
    }
    b.push(0xd2);
    write_leb128_u32(&mut b, e.body_func);
    b.push(super::proposals::stack_switching::OP_CONT_NEW);
    write_leb128_u32(&mut b, e.cont_type);
    b.push(super::proposals::stack_switching::OP_CONT_BIND);
    write_leb128_u32(&mut b, e.cont_type);
    write_leb128_u32(&mut b, ctx.continuation_type_idx);
    local(&mut b, 0x21, cont);
    b.push(0x10);
    write_leb128_u32(&mut b, rt[&("ecma:object", "new")] as u32);
    local(&mut b, 0x21, owner);
    let root = chunks.first().expect("generator module root");
    for (key, alias) in [
        ("next", "__vybe_generator_next"),
        ("iterator", "__vybe_generator_self"),
        ("asyncIterator", "__vybe_generator_self"),
    ] {
        if let Some(original) = root.globals.iter().position(|n| n == alias) {
            local(&mut b, 0x20, owner);
            b.push(0x23);
            write_leb128_u32(&mut b, ctx.string_const_global[key]);
            b.push(0x23);
            write_leb128_u32(&mut b, ctx.callables.global_map[original]);
            b.push(0x10);
            write_leb128_u32(&mut b, rt[&("ecma:object", "set")] as u32);
            b.push(0x1a);
        }
    }
    if e.user_arity > 0 {
        local(&mut b, 0x20, owner);
        b.push(0x23);
        write_leb128_u32(&mut b, ctx.string_const_global["__bound_args"]);
        for i in 1..=e.user_arity {
            local(&mut b, 0x20, u32::from(i));
        }
        typed(&mut b, 8, ctx.array_type_idx);
        write_leb128_u32(&mut b, u32::from(e.user_arity));
        external(&mut b);
        b.push(0x10);
        write_leb128_u32(&mut b, rt[&("ecma:object", "set")] as u32);
        b.push(0x1a);
    }
    local(&mut b, 0x20, cont);
    local(&mut b, 0x20, owner);
    b.push(0xd2);write_leb128_u32(&mut b,ctx.native_resume_function_idx.expect("originating Source driver"));
    gc(&mut b, 0);
    write_leb128_u32(&mut b, ctx.continuation_carrier_type_idx);
    external(&mut b);
    b.push(0x0b);
    b
}

#[cfg(test)]
mod constructor_provenance_tests {
    use super::*;
    use vybe_runtime::chunk::LocalName;
    fn source() -> Chunk {
        let mut c=Chunk::new("unproven");
        c.is_generator=true;c.arity=2;c.local_count=2;c
    }
    #[test]
    fn unmarked_constructor_forms_keep_existing_target() {
        let mut async_form=source();async_form.is_async=true;
        assert!(entries(&[source(),async_form]).is_empty());
        let mut spoof=source();
        spoof.local_names.push(LocalName::new(1,"<native-generator-entry:1:1>",LocalOrigin::Source));
        assert!(entries(&[spoof]).is_empty());
    }
    #[test]
    fn compiler_producer_is_checked_structurally() {
        let mut c=source();
        c.local_names.push(LocalName::new(1,"__generator_entry_control",LocalOrigin::Compiler));
        c.local_names.push(LocalName::new(1,"<native-generator-entry:1:1>",LocalOrigin::Compiler));
        assert_eq!(entries(&[c.clone()]).len(),1);
        c.is_async=true;
        assert!(std::panic::catch_unwind(||entries(&[c])).is_err());
    }
}
