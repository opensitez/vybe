use vybe_compiler::primitives::{class_slots, collections, ops, strings, tuples};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.netip.ParseAddr" if argc == 1 => emit_parse_addr(chunks, current, line),
        "go.netip.MustParseAddr" if argc == 1 => emit_must_parse_addr(chunks, current, line),
        "go.netip.IPv4" if argc == 4 => emit_ipv4(chunks, current, line),
        "go.netip.AddrFromSlice" if argc == 1 => emit_addr_from_slice(chunks, current, line),
        "go.netip.ParsePrefix" if argc == 1 => emit_parse_prefix(chunks, current, line),
        "go.netip.MustParsePrefix" if argc == 1 => emit_must_parse_prefix(chunks, current, line),
        "go.netip.PrefixFrom" if argc == 2 => emit_prefix_from(chunks, current, line),
        "go.netip.ParseAddrPort" if argc == 1 => emit_parse_addr_port(chunks, current, line),
        "go.netip.AddrPortFrom" if argc == 2 => emit_addr_port_from(chunks, current, line),
        "go.netip.Addr.String" if argc == 1 => emit_get_field_stack(chunks, current, "s", line),
        "go.netip.Addr.Is4" if argc == 1 => emit_get_field_stack(chunks, current, "is4", line),
        "go.netip.Addr.Is6" if argc == 1 => emit_get_field_stack(chunks, current, "is6", line),
        "go.netip.Addr.Is4In6" if argc == 1 => {
            emit_get_field_stack(chunks, current, "is4in6", line)
        }
        "go.netip.Addr.IsValid" if argc == 1 => {
            emit_get_field_stack(chunks, current, "valid", line)
        }
        "go.netip.Addr.IsUnspecified" if argc == 1 => {
            emit_addr_predicate(chunks, current, AddrPredicate::Unspecified, line)
        }
        "go.netip.Addr.IsLoopback" if argc == 1 => {
            emit_addr_predicate(chunks, current, AddrPredicate::Loopback, line)
        }
        "go.netip.Addr.IsPrivate" if argc == 1 => {
            emit_addr_predicate(chunks, current, AddrPredicate::Private, line)
        }
        "go.netip.Addr.IsGlobalUnicast" if argc == 1 => {
            emit_is_global_unicast(chunks, current, line)
        }
        "go.netip.Addr.IsLinkLocalUnicast" if argc == 1 => {
            emit_addr_predicate(chunks, current, AddrPredicate::LinkLocal, line)
        }
        "go.netip.Addr.IsMulticast" if argc == 1 => emit_is_multicast(chunks, current, line),
        "go.netip.Addr.Unmap" if argc == 1 => emit_unmap(chunks, current, line),
        "go.netip.Addr.WithZone" if argc == 2 => emit_with_zone(chunks, current, line),
        "go.netip.Addr.Zone" if argc == 1 => emit_get_field_stack(chunks, current, "zone", line),
        "go.netip.Addr.Compare" if argc == 2 => emit_compare(chunks, current, line),
        "go.netip.Addr.Equal" if argc == 2 => emit_equal(chunks, current, line),
        "go.netip.Addr.Less" if argc == 2 => emit_less(chunks, current, line),
        "go.netip.Addr.AsSlice" if argc == 1 => emit_as_slice(chunks, current, line),
        "go.netip.Addr.As16" if argc == 1 => emit_as16(chunks, current, line),
        "go.netip.Addr.Next" | "go.netip.Addr.Prev" if argc == 1 => {}
        "go.netip.Prefix.String" if argc == 1 => emit_prefix_string(chunks, current, line),
        "go.netip.Prefix.Bits" if argc == 1 => emit_get_field_stack(chunks, current, "bits", line),
        "go.netip.Prefix.IsValid" if argc == 1 => {
            emit_get_field_stack(chunks, current, "valid", line)
        }
        "go.netip.Prefix.Addr" if argc == 1 => emit_get_field_stack(chunks, current, "addr", line),
        "go.netip.Prefix.Masked" if argc == 1 => emit_masked(chunks, current, line),
        "go.netip.Prefix.Contains" if argc == 2 => emit_contains(chunks, current, line),
        "go.netip.Prefix.Overlaps" if argc == 2 => emit_overlaps(chunks, current, line),
        "go.netip.Prefix.ContainsPrefix" if argc == 2 => {
            emit_contains_prefix(chunks, current, line)
        }
        "go.netip.AddrPort.String" if argc == 1 => emit_addr_port_string(chunks, current, line),
        "go.netip.AddrPort.Addr" if argc == 1 => {
            emit_get_field_stack(chunks, current, "addr", line)
        }
        "go.netip.AddrPort.Port" if argc == 1 => {
            emit_get_field_stack(chunks, current, "port", line)
        }
        _ => return false,
    }
    true
}

fn emit_parse_addr(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = stash_args(chunks, current, 1, line);
    emit_slot_contains(&mut chunks[current], s, ".", line);
    emit_slot_contains(&mut chunks[current], s, ":", line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    emit_parse_ipv4_tuple(chunks, current, s, line);
    chunks[current].emit_else(line);
    emit_slot_contains(&mut chunks[current], s, ":", line);
    chunks[current].emit_if(line);
    emit_parse_ipv6_tuple(chunks, current, s, line);
    chunks[current].emit_else(line);
    emit_invalid_tuple(chunks, current, NetipKind::Addr, "invalid IP", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_must_parse_addr(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = stash_args(chunks, current, 1, line);
    emit_slot_contains(&mut chunks[current], s, ".", line);
    emit_slot_contains(&mut chunks[current], s, ":", line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    emit_parse_ipv4_value(chunks, current, s, line);
    chunks[current].emit_else(line);
    emit_parse_ipv6_value(chunks, current, s, line);
    chunks[current].emit_end(line);
}

fn emit_ipv4(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 4, line);
    let s = chunks[current].alloc_scratch(1);
    emit_ipv4_string_from_slots(
        &mut chunks[current],
        base,
        base + 1,
        base + 2,
        base + 3,
        line,
    );
    lset(&mut chunks[current], s, line);
    emit_addr_from_slots(
        &mut chunks[current],
        s,
        ConstOrSlot::Bool(true),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
}

fn emit_addr_from_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    let b = stash_args(chunks, current, 1, line);
    lget(&mut chunks[current], b, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    let base = chunks[current].alloc_scratch(4);
    for i in 0..4 {
        lget(&mut chunks[current], b, line);
        chunks[current].emit_i32_const(i, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        lset(&mut chunks[current], base + i as u16, line);
    }
    emit_ipv4_string_from_slots(
        &mut chunks[current],
        base,
        base + 1,
        base + 2,
        base + 3,
        line,
    );
    let s = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], s, line);
    emit_addr_from_slots(
        &mut chunks[current],
        s,
        ConstOrSlot::Bool(true),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
    chunks[current].emit_bool_const(true, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], b, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(16, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    emit_addr_const(
        &mut chunks[current],
        "::",
        false,
        true,
        false,
        "",
        true,
        line,
    );
    chunks[current].emit_bool_const(true, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_else(line);
    emit_zero_addr(&mut chunks[current], line);
    chunks[current].emit_bool_const(false, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_parse_prefix(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = stash_args(chunks, current, 1, line);
    emit_prefix_tuple_from_slot(chunks, current, s, true, line);
}

fn emit_must_parse_prefix(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = stash_args(chunks, current, 1, line);
    emit_prefix_tuple_from_slot(chunks, current, s, false, line);
}

fn emit_prefix_from(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let addr = base;
    let bits = base + 1;
    emit_get_field_local(&mut chunks[current], addr, "is4", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(32.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(128.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    emit_invalid_tuple(chunks, current, NetipKind::Prefix, "invalid prefix", line);
    chunks[current].emit_else(line);
    emit_prefix_from_slots(
        &mut chunks[current],
        addr,
        bits,
        ConstOrSlot::Bool(true),
        line,
    );
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_end(line);
}

fn emit_parse_addr_port(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = stash_args(chunks, current, 1, line);
    emit_slot_starts_with(&mut chunks[current], s, "[", line);
    chunks[current].emit_if(line);
    emit_parse_bracketed_addr_port(chunks, current, s, line);
    chunks[current].emit_else(line);
    emit_parse_plain_addr_port(chunks, current, s, line);
    chunks[current].emit_end(line);
}

fn emit_addr_port_from(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    emit_addr_port_from_slots(
        &mut chunks[current],
        base,
        base + 1,
        ConstOrSlot::Bool(true),
        line,
    );
}

#[derive(Clone, Copy)]
enum AddrPredicate {
    Unspecified,
    Loopback,
    Private,
    LinkLocal,
}

#[derive(Clone, Copy)]
enum NetipKind {
    Addr,
    Prefix,
    AddrPort,
}

#[derive(Clone, Copy)]
enum ConstOrSlot {
    Bool(bool),
    Str(&'static str),
    Slot(u16),
}

fn emit_parse_ipv4_tuple(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    emit_validate_ipv4(chunks, current, s, line);
    chunks[current].emit_if(line);
    emit_parse_ipv4_value(chunks, current, s, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_else(line);
    emit_invalid_tuple(chunks, current, NetipKind::Addr, "invalid IP", line);
    chunks[current].emit_end(line);
}

fn emit_parse_ipv4_value(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    let parts = chunks[current].alloc_scratch(1);
    let nums = chunks[current].alloc_scratch(4);
    split_slot_to(&mut chunks[current], s, ".", parts, line);
    for i in 0..4 {
        parse_part_to(&mut chunks[current], parts, i, nums + i as u16, line);
    }
    let out = chunks[current].alloc_scratch(1);
    emit_ipv4_string_from_slots(
        &mut chunks[current],
        nums,
        nums + 1,
        nums + 2,
        nums + 3,
        line,
    );
    lset(&mut chunks[current], out, line);
    emit_addr_from_slots(
        &mut chunks[current],
        out,
        ConstOrSlot::Bool(true),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
}

fn emit_parse_ipv6_tuple(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    emit_parse_ipv6_value(chunks, current, s, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_parse_ipv6_value(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let z = base;
    let addr_s = base + 1;
    let zone = base + 2;
    let is4in6 = base + 3;
    lget(&mut chunks[current], s, line);
    chunks[current].emit_string_const("%", line);
    call_import(&mut chunks[current], "ecma:string", "indexOf", 2, line);
    lset(&mut chunks[current], z, line);
    lget(&mut chunks[current], z, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if(line);
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::F64(0.0),
        SliceBound::Slot(z),
        addr_s,
        line,
    );
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::SlotPlus(z, 1.0),
        SliceBound::Len(s),
        zone,
        line,
    );
    chunks[current].emit_else(line);
    lget(&mut chunks[current], s, line);
    lset(&mut chunks[current], addr_s, line);
    chunks[current].emit_string_const("", line);
    lset(&mut chunks[current], zone, line);
    chunks[current].emit_end(line);
    emit_slot_contains(&mut chunks[current], addr_s, ".", line);
    lset(&mut chunks[current], is4in6, line);
    emit_addr_from_slots(
        &mut chunks[current],
        addr_s,
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(true),
        ConstOrSlot::Slot(is4in6),
        ConstOrSlot::Slot(zone),
        ConstOrSlot::Bool(true),
        line,
    );
}

fn emit_prefix_tuple_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    s: u16,
    with_error_tuple: bool,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(6);
    let slash = base;
    let addr_s = base + 1;
    let bits_s = base + 2;
    let bits = base + 3;
    let addr = base + 4;
    let is4 = base + 5;

    lget(&mut chunks[current], s, line);
    chunks[current].emit_string_const("/", line);
    call_import(&mut chunks[current], "ecma:string", "indexOf", 2, line);
    lset(&mut chunks[current], slash, line);

    lget(&mut chunks[current], slash, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    if with_error_tuple {
        emit_invalid_tuple(chunks, current, NetipKind::Prefix, "invalid prefix", line);
    } else {
        emit_zero_prefix(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::F64(0.0),
        SliceBound::Slot(slash),
        addr_s,
        line,
    );
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::SlotPlus(slash, 1.0),
        SliceBound::Len(s),
        bits_s,
        line,
    );
    parse_int_to(&mut chunks[current], bits_s, bits, line);
    emit_slot_contains(&mut chunks[current], addr_s, ".", line);
    lset(&mut chunks[current], is4, line);
    emit_addr_from_slots(
        &mut chunks[current],
        addr_s,
        ConstOrSlot::Slot(is4),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
    lset(&mut chunks[current], addr, line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    lget(&mut chunks[current], is4, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(32.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(128.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    if with_error_tuple {
        emit_invalid_tuple(chunks, current, NetipKind::Prefix, "invalid prefix", line);
    } else {
        emit_zero_prefix(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    emit_prefix_from_slots(
        &mut chunks[current],
        addr,
        bits,
        ConstOrSlot::Bool(true),
        line,
    );
    if with_error_tuple {
        chunks[current].emit_ref_null(HT_EXTERN, line);
        tuples::emit_tuple(chunks, current, 2, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_parse_bracketed_addr_port(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let end = base;
    let host = base + 1;
    let port_s = base + 2;
    let port = base + 3;
    let addr = base + 4;
    lget(&mut chunks[current], s, line);
    chunks[current].emit_string_const("]", line);
    call_import(&mut chunks[current], "ecma:string", "indexOf", 2, line);
    lset(&mut chunks[current], end, line);
    lget(&mut chunks[current], end, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    emit_invalid_tuple(
        chunks,
        current,
        NetipKind::AddrPort,
        "invalid addrport",
        line,
    );
    chunks[current].emit_else(line);
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::F64(1.0),
        SliceBound::Slot(end),
        host,
        line,
    );
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::SlotPlus(end, 2.0),
        SliceBound::Len(s),
        port_s,
        line,
    );
    parse_int_to(&mut chunks[current], port_s, port, line);
    emit_addr_from_slots(
        &mut chunks[current],
        host,
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(true),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
    lset(&mut chunks[current], addr, line);
    emit_addr_port_from_slots(
        &mut chunks[current],
        addr,
        port,
        ConstOrSlot::Bool(true),
        line,
    );
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_end(line);
}

fn emit_parse_plain_addr_port(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let sep = base;
    let host = base + 1;
    let port_s = base + 2;
    let port = base + 3;
    let addr = base + 4;
    lget(&mut chunks[current], s, line);
    chunks[current].emit_string_const(":", line);
    call_import(&mut chunks[current], "ecma:string", "lastIndexOf", 2, line);
    lset(&mut chunks[current], sep, line);
    lget(&mut chunks[current], sep, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    emit_invalid_tuple(chunks, current, NetipKind::AddrPort, "missing port", line);
    chunks[current].emit_else(line);
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::F64(0.0),
        SliceBound::Slot(sep),
        host,
        line,
    );
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::SlotPlus(sep, 1.0),
        SliceBound::Len(s),
        port_s,
        line,
    );
    parse_int_to(&mut chunks[current], port_s, port, line);
    emit_addr_from_slots(
        &mut chunks[current],
        host,
        ConstOrSlot::Bool(true),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
    lset(&mut chunks[current], addr, line);
    emit_addr_port_from_slots(
        &mut chunks[current],
        addr,
        port,
        ConstOrSlot::Bool(true),
        line,
    );
    chunks[current].emit_ref_null(HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_end(line);
}

fn emit_addr_predicate(chunks: &mut [Chunk], current: usize, predicate: AddrPredicate, line: u32) {
    let a = stash_args(chunks, current, 1, line);
    let s = chunks[current].alloc_scratch(1);
    emit_get_field_local(&mut chunks[current], a, "s", line);
    lset(&mut chunks[current], s, line);
    match predicate {
        AddrPredicate::Unspecified => {
            emit_slot_eq_const(&mut chunks[current], s, "0.0.0.0", line);
            emit_slot_eq_const(&mut chunks[current], s, "::", line);
            chunks[current].emit_op(Op::I32_OR, line);
        }
        AddrPredicate::Loopback => {
            emit_slot_starts_with(&mut chunks[current], s, "127.", line);
            emit_slot_eq_const(&mut chunks[current], s, "::1", line);
            chunks[current].emit_op(Op::I32_OR, line);
        }
        AddrPredicate::Private => {
            emit_slot_starts_with(&mut chunks[current], s, "10.", line);
            emit_slot_starts_with(&mut chunks[current], s, "192.168.", line);
            chunks[current].emit_op(Op::I32_OR, line);
            emit_slot_starts_with(&mut chunks[current], s, "172.16.", line);
            chunks[current].emit_op(Op::I32_OR, line);
        }
        AddrPredicate::LinkLocal => {
            emit_slot_starts_with(&mut chunks[current], s, "169.254.", line);
            emit_slot_starts_with(&mut chunks[current], s, "fe80:", line);
            chunks[current].emit_op(Op::I32_OR, line);
        }
    }
}

fn emit_is_global_unicast(chunks: &mut [Chunk], current: usize, line: u32) {
    let a = stash_args(chunks, current, 1, line);
    emit_get_field_local(&mut chunks[current], a, "valid", line);
    emit_addr_predicate_from_slot(chunks, current, a, AddrPredicate::Unspecified, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    emit_is_multicast_from_slot(chunks, current, a, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
}

fn emit_addr_predicate_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    a: u16,
    predicate: AddrPredicate,
    line: u32,
) {
    lget(&mut chunks[current], a, line);
    emit_addr_predicate(chunks, current, predicate, line);
}

fn emit_is_multicast(chunks: &mut [Chunk], current: usize, line: u32) {
    let a = stash_args(chunks, current, 1, line);
    emit_is_multicast_from_slot(chunks, current, a, line);
}

fn emit_is_multicast_from_slot(chunks: &mut [Chunk], current: usize, a: u16, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    emit_get_field_local(&mut chunks[current], a, "s", line);
    lset(&mut chunks[current], s, line);
    emit_get_field_local(&mut chunks[current], a, "is6", line);
    chunks[current].emit_if_value(line);
    emit_slot_starts_with(&mut chunks[current], s, "ff", line);
    chunks[current].emit_else(line);
    let mut first = true;
    for prefix in [
        "224.", "225.", "226.", "227.", "228.", "229.", "230.", "231.", "232.", "233.", "234.",
        "235.", "236.", "237.", "238.", "239.",
    ] {
        emit_slot_starts_with(&mut chunks[current], s, prefix, line);
        if !first {
            chunks[current].emit_op(Op::I32_OR, line);
        }
        first = false;
    }
    chunks[current].emit_end(line);
}

fn emit_unmap(chunks: &mut [Chunk], current: usize, line: u32) {
    let a = stash_args(chunks, current, 1, line);
    let s = chunks[current].alloc_scratch(1);
    emit_get_field_local(&mut chunks[current], a, "s", line);
    lset(&mut chunks[current], s, line);
    emit_slot_starts_with(&mut chunks[current], s, "::ffff:", line);
    chunks[current].emit_if(line);
    let mapped = chunks[current].alloc_scratch(1);
    emit_slice_to(
        &mut chunks[current],
        s,
        SliceBound::F64(7.0),
        SliceBound::Len(s),
        mapped,
        line,
    );
    emit_addr_from_slots(
        &mut chunks[current],
        mapped,
        ConstOrSlot::Bool(true),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
    chunks[current].emit_else(line);
    lget(&mut chunks[current], a, line);
    chunks[current].emit_end(line);
}

fn emit_with_zone(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let a = base;
    let zone = base + 1;
    let s = chunks[current].alloc_scratch(4);
    emit_get_field_local(&mut chunks[current], a, "s", line);
    lset(&mut chunks[current], s, line);
    emit_get_field_local(&mut chunks[current], a, "is4", line);
    lset(&mut chunks[current], s + 1, line);
    emit_get_field_local(&mut chunks[current], a, "is6", line);
    lset(&mut chunks[current], s + 2, line);
    emit_get_field_local(&mut chunks[current], a, "is4in6", line);
    lset(&mut chunks[current], s + 3, line);
    emit_addr_from_slots(
        &mut chunks[current],
        s,
        ConstOrSlot::Slot(s + 1),
        ConstOrSlot::Slot(s + 2),
        ConstOrSlot::Slot(s + 3),
        ConstOrSlot::Slot(zone),
        ConstOrSlot::Bool(true),
        line,
    );
}

fn emit_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    emit_addr_string_to(&mut chunks[current], base, base + 2, line);
    emit_addr_string_to(&mut chunks[current], base + 1, base + 3, line);
    emit_slots_eq(&mut chunks[current], base + 2, base + 3, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_end(line);
}

fn emit_equal(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    emit_addr_string_to(&mut chunks[current], base, base + 2, line);
    emit_addr_string_to(&mut chunks[current], base + 1, base + 3, line);
    emit_slots_eq(&mut chunks[current], base + 2, base + 3, line);
}

fn emit_less(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_compare(chunks, current, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
}

fn emit_as_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    let a = stash_args(chunks, current, 1, line);
    let s = chunks[current].alloc_scratch(1);
    emit_get_field_local(&mut chunks[current], a, "s", line);
    lset(&mut chunks[current], s, line);
    emit_validate_ipv4(chunks, current, s, line);
    chunks[current].emit_if(line);
    emit_ipv4_byte_array(chunks, current, s, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
}

fn emit_as16(chunks: &mut [Chunk], current: usize, _line: u32) {
    let line = _line;
    let _ = stash_args(chunks, current, 1, line);
    for _ in 0..16 {
        chunks[current].emit_f64_const(0.0, line);
    }
    collections::emit_array_new(chunks, current, 16, line);
}

fn emit_prefix_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let p = stash_args(chunks, current, 1, line);
    let base = chunks[current].alloc_scratch(2);
    emit_get_field_local(&mut chunks[current], p, "addr", line);
    lset(&mut chunks[current], base, line);
    emit_addr_string_to(&mut chunks[current], base, base + 1, line);
    lget(&mut chunks[current], base + 1, line);
    chunks[current].emit_string_const("/", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    emit_get_field_local(&mut chunks[current], p, "bits", line);
    call_import(&mut chunks[current], "ecma:string", "String", 1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
}

fn emit_masked(chunks: &mut [Chunk], current: usize, line: u32) {
    let p = stash_args(chunks, current, 1, line);
    let base = chunks[current].alloc_scratch(4);
    let addr = base;
    let bits = base + 1;
    let masked_s = base + 2;
    let is4 = base + 3;
    emit_get_field_local(&mut chunks[current], p, "addr", line);
    lset(&mut chunks[current], addr, line);
    emit_get_field_local(&mut chunks[current], p, "bits", line);
    lset(&mut chunks[current], bits, line);
    emit_get_field_local(&mut chunks[current], addr, "is4", line);
    lset(&mut chunks[current], is4, line);
    lget(&mut chunks[current], is4, line);
    chunks[current].emit_if(line);
    emit_ipv4_mask_to(chunks, current, addr, bits, masked_s, line);
    chunks[current].emit_else(line);
    emit_addr_string_to(&mut chunks[current], addr, masked_s, line);
    chunks[current].emit_end(line);
    emit_addr_from_slots(
        &mut chunks[current],
        masked_s,
        ConstOrSlot::Slot(is4),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Bool(false),
        ConstOrSlot::Str(""),
        ConstOrSlot::Bool(true),
        line,
    );
    lset(&mut chunks[current], addr, line);
    emit_prefix_from_slots(
        &mut chunks[current],
        addr,
        bits,
        ConstOrSlot::Bool(true),
        line,
    );
}

fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    emit_contains_from_slots(chunks, current, base, base + 1, line);
}

fn emit_overlaps(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let q_addr = chunks[current].alloc_scratch(1);
    let p_addr = chunks[current].alloc_scratch(1);
    emit_get_field_local(&mut chunks[current], base + 1, "addr", line);
    lset(&mut chunks[current], q_addr, line);
    emit_contains_from_slots(chunks, current, base, q_addr, line);
    emit_get_field_local(&mut chunks[current], base, "addr", line);
    lset(&mut chunks[current], p_addr, line);
    emit_contains_from_slots(chunks, current, base + 1, p_addr, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

fn emit_contains_prefix(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let q_addr = chunks[current].alloc_scratch(1);
    emit_get_field_local(&mut chunks[current], base + 1, "addr", line);
    lset(&mut chunks[current], q_addr, line);
    emit_contains_from_slots(chunks, current, base, q_addr, line);
    emit_get_field_local(&mut chunks[current], base, "bits", line);
    emit_get_field_local(&mut chunks[current], base + 1, "bits", line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
}

fn emit_addr_port_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let ap = stash_args(chunks, current, 1, line);
    let base = chunks[current].alloc_scratch(3);
    let addr = base;
    let s = base + 1;
    let port = base + 2;
    emit_get_field_local(&mut chunks[current], ap, "addr", line);
    lset(&mut chunks[current], addr, line);
    emit_addr_string_to(&mut chunks[current], addr, s, line);
    emit_get_field_local(&mut chunks[current], ap, "port", line);
    lset(&mut chunks[current], port, line);
    emit_get_field_local(&mut chunks[current], addr, "is6", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("[", line);
    lget(&mut chunks[current], s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const("]:", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], s, line);
    chunks[current].emit_string_const(":", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], port, line);
    call_import(&mut chunks[current], "ecma:string", "String", 1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
}

fn emit_contains_from_slots(chunks: &mut [Chunk], current: usize, p: u16, a: u16, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let p_addr = base;
    let p_s = base + 1;
    let a_s = base + 2;
    let bits = base + 3;
    let first = base + 4;
    emit_get_field_local(&mut chunks[current], p, "addr", line);
    lset(&mut chunks[current], p_addr, line);
    emit_addr_string_to(&mut chunks[current], p_addr, p_s, line);
    emit_addr_string_to(&mut chunks[current], a, a_s, line);
    emit_get_field_local(&mut chunks[current], p, "bits", line);
    lset(&mut chunks[current], bits, line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(8.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    emit_first_ipv4_part_to(&mut chunks[current], p_s, first, line);
    lget(&mut chunks[current], first, line);
    chunks[current].emit_string_const(".", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    let prefix = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], prefix, line);
    emit_slot_starts_with(&mut chunks[current], a_s, "", line);
    lget(&mut chunks[current], a_s, line);
    lget(&mut chunks[current], prefix, line);
    call_import(&mut chunks[current], "ecma:string", "startsWith", 2, line);
    chunks[current].emit_else(line);
    emit_slots_eq(&mut chunks[current], p_s, a_s, line);
    chunks[current].emit_end(line);
}

fn emit_ipv4_mask_to(
    chunks: &mut [Chunk],
    current: usize,
    addr: u16,
    bits: u16,
    dest: u16,
    line: u32,
) {
    let s = chunks[current].alloc_scratch(1);
    emit_addr_string_to(&mut chunks[current], addr, s, line);
    let parts = chunks[current].alloc_scratch(1);
    split_slot_to(&mut chunks[current], s, ".", parts, line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(24.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(16.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_f64_const(8.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    get_part(&mut chunks[current], parts, 0, line);
    chunks[current].emit_string_const(".0.0.0", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get_part(&mut chunks[current], parts, 0, line);
    chunks[current].emit_string_const(".", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    get_part(&mut chunks[current], parts, 1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(".0.0", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    get_part(&mut chunks[current], parts, 0, line);
    chunks[current].emit_string_const(".", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    get_part(&mut chunks[current], parts, 1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(".", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    get_part(&mut chunks[current], parts, 2, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(".0", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], s, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], dest, line);
}

fn emit_validate_ipv4(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    let parts = chunks[current].alloc_scratch(1);
    let nums = chunks[current].alloc_scratch(4);
    split_slot_to(&mut chunks[current], s, ".", parts, line);
    lget(&mut chunks[current], parts, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    for i in 0..4 {
        parse_part_to(&mut chunks[current], parts, i, nums + i as u16, line);
        emit_part_valid(&mut chunks[current], parts, i, nums + i as u16, line);
        chunks[current].emit_op(Op::I32_AND, line);
    }
}

fn emit_ipv4_byte_array(chunks: &mut [Chunk], current: usize, s: u16, line: u32) {
    let parts = chunks[current].alloc_scratch(1);
    split_slot_to(&mut chunks[current], s, ".", parts, line);
    for i in 0..4 {
        get_part(&mut chunks[current], parts, i, line);
        call_import(&mut chunks[current], "ecma:number", "parseInt", 1, line);
    }
    collections::emit_array_new(chunks, current, 4, line);
}

fn emit_ipv4_string_from_slots(chunk: &mut Chunk, a: u16, b: u16, c: u16, d: u16, line: u32) {
    lget(chunk, a, line);
    call_import(chunk, "ecma:string", "String", 1, line);
    chunk.emit_string_const(".", line);
    ops::emit_dyn_add(chunk, line);
    lget(chunk, b, line);
    call_import(chunk, "ecma:string", "String", 1, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_string_const(".", line);
    ops::emit_dyn_add(chunk, line);
    lget(chunk, c, line);
    call_import(chunk, "ecma:string", "String", 1, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_string_const(".", line);
    ops::emit_dyn_add(chunk, line);
    lget(chunk, d, line);
    call_import(chunk, "ecma:string", "String", 1, line);
    ops::emit_dyn_add(chunk, line);
}

fn emit_addr_string_to(chunk: &mut Chunk, addr: u16, dest: u16, line: u32) {
    emit_get_field_local(chunk, addr, "s", line);
    lset(chunk, dest, line);
}

fn emit_first_ipv4_part_to(chunk: &mut Chunk, s: u16, dest: u16, line: u32) {
    let parts = chunk.alloc_scratch(1);
    split_slot_to(chunk, s, ".", parts, line);
    get_part(chunk, parts, 0, line);
    lset(chunk, dest, line);
}

fn parse_part_to(chunk: &mut Chunk, parts: u16, index: i32, dest: u16, line: u32) {
    get_part(chunk, parts, index, line);
    call_import(chunk, "ecma:number", "parseInt", 1, line);
    lset(chunk, dest, line);
}

fn parse_int_to(chunk: &mut Chunk, s: u16, dest: u16, line: u32) {
    lget(chunk, s, line);
    call_import(chunk, "ecma:number", "parseInt", 1, line);
    lset(chunk, dest, line);
}

fn emit_part_valid(chunk: &mut Chunk, parts: u16, index: i32, value: u16, line: u32) {
    get_part(chunk, parts, index, line);
    strings::emit_length(chunk, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    lget(chunk, value, line);
    lget(chunk, value, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, value, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, value, line);
    chunk.emit_f64_const(255.0, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn split_slot_to(chunk: &mut Chunk, s: u16, sep: &str, dest: u16, line: u32) {
    lget(chunk, s, line);
    chunk.emit_string_const(sep, line);
    call_import(chunk, "ecma:string", "split", 2, line);
    lset(chunk, dest, line);
}

fn get_part(chunk: &mut Chunk, parts: u16, index: i32, line: u32) {
    lget(chunk, parts, line);
    chunk.emit_i32_const(index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_slot_contains(chunk: &mut Chunk, slot: u16, needle: &str, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_string_const(needle, line);
    call_import(chunk, "ecma:string", "includes", 2, line);
}

fn emit_slot_starts_with(chunk: &mut Chunk, slot: u16, prefix: &str, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_string_const(prefix, line);
    call_import(chunk, "ecma:string", "startsWith", 2, line);
}

fn emit_slot_eq_const(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_string_const(value, line);
    ops::emit_dyn_eq(chunk, line);
}

fn emit_slots_eq(chunk: &mut Chunk, left: u16, right: u16, line: u32) {
    lget(chunk, left, line);
    lget(chunk, right, line);
    ops::emit_dyn_eq(chunk, line);
}

enum SliceBound {
    F64(f64),
    Slot(u16),
    SlotPlus(u16, f64),
    Len(u16),
}

fn emit_slice_to(
    chunk: &mut Chunk,
    s: u16,
    start: SliceBound,
    end: SliceBound,
    dest: u16,
    line: u32,
) {
    lget(chunk, s, line);
    emit_bound(chunk, start, line);
    emit_bound(chunk, end, line);
    call_import(chunk, "ecma:string", "slice", 3, line);
    lset(chunk, dest, line);
}

fn emit_bound(chunk: &mut Chunk, bound: SliceBound, line: u32) {
    match bound {
        SliceBound::F64(value) => chunk.emit_f64_const(value, line),
        SliceBound::Slot(slot) => lget(chunk, slot, line),
        SliceBound::SlotPlus(slot, offset) => {
            lget(chunk, slot, line);
            chunk.emit_f64_const(offset, line);
            chunk.emit_op(Op::F64_ADD, line);
        }
        SliceBound::Len(slot) => {
            lget(chunk, slot, line);
            strings::emit_length(chunk, line);
        }
    }
}

fn emit_invalid_tuple(
    chunks: &mut [Chunk],
    current: usize,
    kind: NetipKind,
    message: &str,
    line: u32,
) {
    match kind {
        NetipKind::Addr => emit_zero_addr(&mut chunks[current], line),
        NetipKind::Prefix => emit_zero_prefix(&mut chunks[current], line),
        NetipKind::AddrPort => emit_zero_addr_port(&mut chunks[current], line),
    }
    chunks[current].emit_string_const(message, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_zero_addr(chunk: &mut Chunk, line: u32) {
    emit_addr_const(chunk, "", false, false, false, "", false, line);
}

fn emit_zero_prefix(chunk: &mut Chunk, line: u32) {
    let addr = chunk.alloc_scratch(1);
    emit_zero_addr(chunk, line);
    lset(chunk, addr, line);
    emit_prefix_from_slots(chunk, addr, addr, ConstOrSlot::Bool(false), line);
}

fn emit_zero_addr_port(chunk: &mut Chunk, line: u32) {
    let addr = chunk.alloc_scratch(1);
    emit_zero_addr(chunk, line);
    lset(chunk, addr, line);
    emit_addr_port_from_slots(chunk, addr, addr, ConstOrSlot::Bool(false), line);
}

fn emit_addr_const(
    chunk: &mut Chunk,
    s: &'static str,
    is4: bool,
    is6: bool,
    is4in6: bool,
    zone: &'static str,
    valid: bool,
    line: u32,
) {
    chunk.emit_string_const(s, line);
    let s_slot = chunk.alloc_scratch(1);
    lset(chunk, s_slot, line);
    emit_addr_from_slots(
        chunk,
        s_slot,
        ConstOrSlot::Bool(is4),
        ConstOrSlot::Bool(is6),
        ConstOrSlot::Bool(is4in6),
        ConstOrSlot::Str(zone),
        ConstOrSlot::Bool(valid),
        line,
    );
}

fn emit_addr_from_slots(
    chunk: &mut Chunk,
    s: u16,
    is4: ConstOrSlot,
    is6: ConstOrSlot,
    is4in6: ConstOrSlot,
    zone: ConstOrSlot,
    valid: ConstOrSlot,
    line: u32,
) {
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    lset(chunk, obj, line);
    set_local(chunk, obj, "s", s, line);
    set_value(chunk, obj, "is4", is4, line);
    set_value(chunk, obj, "is6", is6, line);
    set_value(chunk, obj, "is4in6", is4in6, line);
    set_value(chunk, obj, "zone", zone, line);
    set_value(chunk, obj, "valid", valid, line);
    set_const_str(chunk, obj, "__type", "__goNetipAddr", line);
    lget(chunk, obj, line);
}

fn emit_prefix_from_slots(chunk: &mut Chunk, addr: u16, bits: u16, valid: ConstOrSlot, line: u32) {
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    lset(chunk, obj, line);
    set_local(chunk, obj, "addr", addr, line);
    set_local(chunk, obj, "bits", bits, line);
    set_value(chunk, obj, "valid", valid, line);
    set_const_str(chunk, obj, "__type", "__goNetipPrefix", line);
    lget(chunk, obj, line);
}

fn emit_addr_port_from_slots(
    chunk: &mut Chunk,
    addr: u16,
    port: u16,
    valid: ConstOrSlot,
    line: u32,
) {
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    lset(chunk, obj, line);
    set_local(chunk, obj, "addr", addr, line);
    set_local(chunk, obj, "port", port, line);
    set_value(chunk, obj, "valid", valid, line);
    set_const_str(chunk, obj, "__type", "__goNetipAddrPort", line);
    lget(chunk, obj, line);
}

fn emit_get_field_stack(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let obj = stash_args(chunks, current, 1, line);
    emit_get_field_local(&mut chunks[current], obj, field, line);
}

fn emit_get_field_local(chunk: &mut Chunk, obj: u16, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot(field),
        class_slots::Dest::Stack,
        line,
    );
}

fn set_local(chunk: &mut Chunk, obj: u16, field: &str, value: u16, line: u32) {
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot(field),
        class_slots::ValueSource::Local(value),
        line,
    );
}

fn set_value(chunk: &mut Chunk, obj: u16, field: &str, value: ConstOrSlot, line: u32) {
    match value {
        ConstOrSlot::Bool(value) => class_slots::emit_class_set(
            chunk,
            class_slots::ObjSource::Local(obj),
            &slot(field),
            class_slots::ValueSource::ConstBool(value),
            line,
        ),
        ConstOrSlot::Str(value) => set_const_str(chunk, obj, field, value, line),
        ConstOrSlot::Slot(value) => set_local(chunk, obj, field, value, line),
    }
}

fn set_const_str(chunk: &mut Chunk, obj: u16, field: &str, value: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot(field),
        class_slots::ValueSource::ConstStr(value.to_string()),
        line,
    );
}

fn slot(field: &str) -> class_slots::ResolvedSlot {
    class_slots::resolve(
        &class_slots::ClassSlot::internal(field),
        &class_slots::PlainNames,
    )
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}
