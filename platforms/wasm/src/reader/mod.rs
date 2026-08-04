//! WASM binary reader — decodes .wasm files into Chunk arrays.

use crate::encoding::*;
use std::collections::HashSet;
use std::sync::Arc;
use vybe_runtime::chunk::{ActiveDataSegment, ActiveElementSegment, StackSwitchHandler};
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op};

#[derive(Default)]
struct StandardSections {
    type_section: Vec<u8>,
    import_section: Vec<u8>,
    func_section: Vec<u8>,
    table_section: Vec<u8>,
    memory_section: Vec<u8>,
    global_section: Vec<u8>,
    export_section: Vec<u8>,
    start_section: Vec<u8>,
    elem_section: Vec<u8>,
    code_section: Vec<u8>,
    data_section: Vec<u8>,
    data_count_section: Vec<u8>,
    tag_section: Vec<u8> }

pub fn read_wasm(data: &[u8]) -> Result<Vec<Chunk>, String> {
    if data.len() < 8 || &data[0..4] != &WASM_MAGIC {
        return Err("Invalid WASM: bad magic".into());
    }
    let mut pos = 8;
    let mut custom_data: Option<Vec<u8>> = None;
    let mut sections = StandardSections::default();
    let mut seen_sections = HashSet::new();
    let mut last_known_section = 0u8;

    while pos < data.len() {
        let section_id = data[pos];
        pos += 1;
        let (size, read) = read_leb128_u32(&data[pos..]);
        if read == 0 {
            return Err("Invalid WASM: malformed section size".into());
        }
        pos += read;
        let section_end = pos
            .checked_add(size as usize)
            .ok_or_else(|| "Invalid WASM: section size overflow".to_string())?;
        if section_end > data.len() {
            return Err("Invalid WASM: truncated section payload".into());
        }
        let section_data = data[pos..section_end].to_vec();

        if section_id != SECTION_CUSTOM {
            if !seen_sections.insert(section_id) {
                return Err(format!("Invalid WASM: duplicate section {section_id}"));
            }
            let rank = section_order_rank(section_id);
            if rank < section_order_rank(last_known_section) {
                return Err(format!("Invalid WASM: section {section_id} out of order"));
            }
            last_known_section = section_id;
        }

        match section_id {
            SECTION_CUSTOM => {
                // Check if it's our "vybe" custom section
                let (nlen, nr) = read_leb128_u32(&section_data);
                if nlen == 4 && section_data.get(nr..nr + 4) == Some(b"vybe") {
                    custom_data = Some(section_data);
                }
            }
            SECTION_TYPE => sections.type_section = section_data,
            SECTION_IMPORT => sections.import_section = section_data,
            SECTION_FUNCTION => sections.func_section = section_data,
            4 => sections.table_section = section_data,
            SECTION_MEMORY => sections.memory_section = section_data,
            SECTION_GLOBAL => sections.global_section = section_data,
            SECTION_EXPORT => sections.export_section = section_data,
            8 => sections.start_section = section_data,
            9 => sections.elem_section = section_data,
            SECTION_CODE => sections.code_section = section_data,
            11 => sections.data_section = section_data,
            12 => sections.data_count_section = section_data,
            SECTION_TAG => sections.tag_section = section_data,
            _ => {}
        }
        pos = section_end;
    }

    // If we have a vybe custom section, use that for round-trip (our format)
    if let Some(ref cd) = custom_data {
        return decode_vybe_section(cd);
    }

    // Otherwise, decode as standard WASM module
    if sections.code_section.is_empty() {
        return Err("No code section in WASM module".into());
    }
    validate_standard_sections(&sections)?;
    decode_standard_wasm(
        &sections.type_section,
        &sections.import_section,
        &sections.func_section,
        &sections.table_section,
        &sections.memory_section,
        &sections.export_section,
        &sections.elem_section,
        &sections.code_section,
        &sections.data_section,
        &sections.tag_section,
    )
}

/// Rejection message for the legacy (pre-3.0) exception-handling proposal.
/// We import the standardized `try_table`/`throw`/`throw_ref` model (WASM 3.0);
/// the legacy `try`/`catch`/`catch_all`/`delegate`/`rethrow` opcodes are not
/// supported (same stance as Wasmtime). Reject them loudly rather than
/// silently mis-decode.
const LEGACY_EH_UNSUPPORTED: &str = "Unsupported WASM: legacy exception-handling \
     (pre-3.0 try/catch/catch_all/delegate/rethrow) is not supported; recompile \
     with the standard exception-handling model (try_table)";

fn section_order_rank(section_id: u8) -> u8 {
    match section_id {
        SECTION_CUSTOM => 0,
        SECTION_TYPE => 1,
        SECTION_IMPORT => 2,
        SECTION_FUNCTION => 3,
        4 => 4, // table
        SECTION_MEMORY => 5,
        // Tag section (exception-handling proposal): ordered after memory,
        // before global — where real toolchains place it.
        SECTION_TAG => 6,
        SECTION_GLOBAL => 7,
        SECTION_EXPORT => 8,
        8 => 9,   // start
        9 => 10,  // element
        12 => 11, // data_count is ordered before code
        SECTION_CODE => 12,
        11 => 13, // data
        other => other }
}

fn validate_standard_sections(sections: &StandardSections) -> Result<(), String> {
    let types = parse_type_section(&sections.type_section);
    let func_type_indices = parse_function_section(&sections.func_section);
    let imports = parse_import_details(&sections.import_section)?;
    for import in &imports {
        if import.kind == 0 && import.type_index as usize >= types.len() {
            return Err("Invalid WASM: import function type index out of range".into());
        }
    }

    for &type_idx in &func_type_indices {
        if type_idx as usize >= types.len() {
            return Err("Invalid WASM: function type index out of range".into());
        }
    }

    let code_count = section_count(&sections.code_section)?;
    if code_count as usize != func_type_indices.len() {
        return Err("Invalid WASM: function/code count mismatch".into());
    }

    validate_memory_section(&sections.memory_section)?;

    let import_func_count = imports.iter().filter(|import| import.kind == 0).count();
    let import_table_count = imports.iter().filter(|import| import.kind == 1).count();
    let import_memory_count = imports.iter().filter(|import| import.kind == 2).count();
    let import_global_count = imports.iter().filter(|import| import.kind == 3).count();

    let table_count = import_table_count + section_count_or_zero(&sections.table_section)? as usize;
    let memory_count =
        import_memory_count + section_count_or_zero(&sections.memory_section)? as usize;
    let (global_count, global_mutability) =
        parse_global_mutability(&sections.global_section, import_global_count)?;
    let func_count = import_func_count + func_type_indices.len();

    validate_exports(&sections.export_section, func_count)?;
    validate_start(
        &sections.start_section,
        &types,
        &imports,
        &func_type_indices,
    )?;
    validate_element_section(&sections.elem_section, table_count)?;
    validate_data_sections(
        &sections.data_count_section,
        &sections.data_section,
        memory_count,
    )?;
    // Full function signature list (imports first, then local funcs) —
    // arities feed the spec validation algorithm's call effects.
    let mut func_sigs: Vec<(usize, usize)> = Vec::with_capacity(func_count);
    for import in &imports {
        if import.kind == 0 {
            let (params, results) = &types[import.type_index as usize];
            func_sigs.push((params.len(), results.len()));
        }
    }
    for &type_idx in &func_type_indices {
        let (params, results) = &types[type_idx as usize];
        func_sigs.push((params.len(), results.len()));
    }

    validate_code_bodies(
        &sections.code_section,
        &func_type_indices,
        &types,
        &func_sigs,
        global_count,
        &global_mutability,
        section_count_or_zero(&sections.data_section)? as usize,
        section_count_or_zero(&sections.elem_section)? as usize,
        !sections.data_count_section.is_empty(),
        section_uses_memory64(&sections.memory_section),
        section_uses_table64(&sections.table_section),
    )
}

struct ImportDetail {
    kind: u8,
    type_index: u32 }

/// Custom Descriptors adds `externtype ::= ... | 0x20 x:typeidx => func exact x`
/// — the same payload as an ordinary function import (`0x00`), refined so a
/// `ref.func` on it may be typed exactly. The refinement is a validation-time
/// property with no representation in our bytecode, so an exact function
/// import is normalised to kind 0 here. That matters: every `kind == 0` site
/// counts imported functions, and an unrecognised kind would shift the whole
/// function index space.
fn normalize_import_kind(kind: u8) -> u8 {
    if kind == EXTERNTYPE_FUNC_EXACT { 0 } else { kind }
}

fn skip_import_descriptor(data: &[u8], pos: &mut usize, kind: u8) {
    match normalize_import_kind(kind) {
        0 => skip_leb128(data, pos), // type index
        1 => {
            skip_leb128(data, pos); // reftype
            let _ = read_limits_min(data, pos);
        }
        2 => {
            let _ = read_limits_min(data, pos);
        }
        3 => {
            skip_leb128(data, pos); // valtype
            *pos = (*pos).saturating_add(1).min(data.len()); // mutability
        }
        _ => {}
    }
}

fn parse_import_details(data: &[u8]) -> Result<Vec<ImportDetail>, String> {
    if data.is_empty() {
        return Ok(Vec::new());
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut imports = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + nlen as usize;
        if pos >= data.len() {
            return Err("Invalid WASM: malformed import section".into());
        }
        let kind = normalize_import_kind(data[pos]);
        pos += 1;
        let type_index = if kind == 0 {
            let (type_index, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            type_index
        } else {
            skip_import_descriptor(data, &mut pos, kind);
            0
        };
        imports.push(ImportDetail { kind, type_index });
    }
    Ok(imports)
}

fn section_count(data: &[u8]) -> Result<u32, String> {
    if data.is_empty() {
        return Err("Invalid WASM: missing required section count".into());
    }
    let (count, read) = read_leb128_u32(data);
    if read == 0 {
        return Err("Invalid WASM: malformed section count".into());
    }
    Ok(count)
}

fn section_count_or_zero(data: &[u8]) -> Result<u32, String> {
    if data.is_empty() {
        Ok(0)
    } else {
        section_count(data)
    }
}

fn validate_memory_section(data: &[u8]) -> Result<(), String> {
    if data.is_empty() {
        return Ok(());
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    for _ in 0..count {
        if pos >= data.len() {
            return Err("Invalid WASM: malformed memory section".into());
        }
        let flags = data[pos];
        pos += 1;
        let is_memory64 = flags & 0x04 != 0;
        let has_max = flags & 0x01 != 0;
        let (min, read) = if is_memory64 {
            read_leb128_u64_local(&data[pos..])
        } else {
            let (value, read) = read_leb128_u32(&data[pos..]);
            (value as u64, read)
        };
        pos += read;
        if has_max {
            let (max, read) = if is_memory64 {
                read_leb128_u64_local(&data[pos..])
            } else {
                let (value, read) = read_leb128_u32(&data[pos..]);
                (value as u64, read)
            };
            pos += read;
            if min > max {
                return Err("Invalid WASM: memory minimum exceeds maximum".into());
            }
        }
    }
    Ok(())
}

fn parse_global_mutability(
    data: &[u8],
    imported_globals: usize,
) -> Result<(usize, Vec<bool>), String> {
    let mut mutability = vec![true; imported_globals];
    if data.is_empty() {
        return Ok((imported_globals, mutability));
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    for _ in 0..count {
        if pos + 2 > data.len() {
            return Err("Invalid WASM: malformed global section".into());
        }
        pos += 1; // valtype
        let mutable = data[pos] != 0;
        pos += 1;
        mutability.push(mutable);
        while pos < data.len() {
            let op = data[pos];
            pos += 1;
            match op {
                0x0B => break,
                0x41 => skip_leb128(data, &mut pos),
                0x42 => skip_leb128(data, &mut pos),
                0x43 => pos += 4,
                0x44 => pos += 8,
                0x23 => skip_leb128(data, &mut pos),
                0xD0 => pos += 1,
                0xD2 => skip_leb128(data, &mut pos),
                _ => {}
            }
        }
    }
    Ok((mutability.len(), mutability))
}

fn validate_exports(data: &[u8], func_count: usize) -> Result<(), String> {
    if data.is_empty() {
        return Ok(());
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut names = HashSet::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name_end = pos + nlen as usize;
        if name_end > data.len() {
            return Err("Invalid WASM: malformed export name".into());
        }
        let name = &data[pos..name_end];
        pos = name_end;
        if !names.insert(name.to_vec()) {
            return Err("Invalid WASM: duplicate export name".into());
        }
        if pos >= data.len() {
            return Err("Invalid WASM: malformed export section".into());
        }
        let kind = data[pos];
        pos += 1;
        // Exports never declare exactness: an exported function is exact iff
        // its internal type is. Custom Descriptors states outright that an
        // export section using 0x20 is malformed.
        if kind == EXTERNTYPE_FUNC_EXACT {
            return Err("Invalid WASM: exact function type in export section".into());
        }
        let (idx, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        if kind == 0 && idx as usize >= func_count {
            return Err("Invalid WASM: function export index out of range".into());
        }
    }
    Ok(())
}

fn validate_start(
    data: &[u8],
    types: &[(Vec<u8>, Vec<u8>)],
    imports: &[ImportDetail],
    func_type_indices: &[u32],
) -> Result<(), String> {
    if data.is_empty() {
        return Ok(());
    }
    let (idx, read) = read_leb128_u32(data);
    if read == 0 || read != data.len() {
        return Err("Invalid WASM: malformed start section".into());
    }
    let import_func_count = imports.iter().filter(|import| import.kind == 0).count();
    let type_idx = if (idx as usize) < import_func_count {
        imports
            .iter()
            .filter(|import| import.kind == 0)
            .nth(idx as usize)
            .map(|import| import.type_index)
    } else {
        func_type_indices
            .get(idx as usize - import_func_count)
            .copied()
    };
    let Some(type_idx) = type_idx else {
        return Err("Invalid WASM: start function index out of range".into());
    };
    let Some((params, results)) = types.get(type_idx as usize) else {
        return Err("Invalid WASM: start function type index out of range".into());
    };
    if !params.is_empty() || !results.is_empty() {
        return Err("Invalid WASM: start function must have type [] -> []".into());
    }
    Ok(())
}

fn validate_element_section(data: &[u8], table_count: usize) -> Result<(), String> {
    if data.is_empty() {
        return Ok(());
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    for _ in 0..count {
        let (flags, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        if flags == 2 {
            let (table_idx, _) = read_leb128_u32(&data[pos..]);
            if table_idx as usize >= table_count {
                return Err("Invalid WASM: element segment table index out of range".into());
            }
        } else if flags == 0 && table_count == 0 {
            return Err("Invalid WASM: active element segment without table".into());
        }
        // Full element-section validation is larger; this pass covers active
        // table-index validity for the runtime reader.
        break;
    }
    Ok(())
}

fn validate_data_sections(
    data_count_section: &[u8],
    data_section: &[u8],
    memory_count: usize,
) -> Result<(), String> {
    let actual_count = section_count_or_zero(data_section)?;
    if !data_count_section.is_empty() {
        let declared = section_count(data_count_section)?;
        if declared != actual_count {
            return Err("Invalid WASM: data_count does not match data section".into());
        }
    }
    if data_section.is_empty() {
        return Ok(());
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data_section[pos..]);
    pos += read;
    for _ in 0..count {
        let (flags, read) = read_leb128_u32(&data_section[pos..]);
        pos += read;
        if flags == 2 {
            let (memidx, _) = read_leb128_u32(&data_section[pos..]);
            if memidx as usize >= memory_count {
                return Err("Invalid WASM: data segment memory index out of range".into());
            }
        } else if flags == 0 && memory_count == 0 {
            return Err("Invalid WASM: active data segment without memory".into());
        }
        break;
    }
    Ok(())
}

fn parse_data_segments(data: &[u8]) -> Result<(Vec<Vec<u8>>, Vec<ActiveDataSegment>), String> {
    if data.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut segments = Vec::with_capacity(count as usize);
    let mut active = Vec::new();
    for data_index in 0..count {
        let (flags, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut active_init = None;
        match flags {
            0 => {
                let offset = read_i32_const_expr_as_u64(data, &mut pos)?;
                active_init = Some((0, offset));
            }
            1 => {}
            2 => {
                let (memidx, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let offset = read_i32_const_expr_as_u64(data, &mut pos)?;
                active_init = Some((memidx, offset));
            }
            _ => return Err("Invalid WASM: unsupported data segment mode".into()) }
        let (len, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let end = pos
            .checked_add(len as usize)
            .ok_or_else(|| "Invalid WASM: data segment size overflow".to_string())?;
        if end > data.len() {
            return Err("Invalid WASM: truncated data segment".into());
        }
        segments.push(data[pos..end].to_vec());
        if let Some((memory_index, offset)) = active_init {
            active.push(ActiveDataSegment {
                memory_index,
                offset,
                data_index });
        }
        pos = end;
    }
    Ok((segments, active))
}

fn parse_element_segments(
    data: &[u8],
) -> Result<(Vec<Vec<Value>>, Vec<ActiveElementSegment>), String> {
    if data.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut segments = Vec::with_capacity(count as usize);
    let mut active = Vec::new();
    for elem_index in 0..count {
        let (flags, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut active_init = None;
        let expr_items = match flags {
            0 => {
                let offset = read_i32_const_expr_as_u64(data, &mut pos)?;
                active_init = Some((0, offset));
                false
            }
            1 => {
                pos += 1; // elemkind
                false
            }
            2 => {
                let (tableidx, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let offset = read_i32_const_expr_as_u64(data, &mut pos)?;
                active_init = Some((tableidx, offset));
                pos += 1; // elemkind
                false
            }
            3 => {
                pos += 1; // elemkind
                false
            }
            4 => {
                let offset = read_i32_const_expr_as_u64(data, &mut pos)?;
                active_init = Some((0, offset));
                true
            }
            5 => {
                skip_leb128(data, &mut pos); // reftype
                true
            }
            6 => {
                let (tableidx, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let offset = read_i32_const_expr_as_u64(data, &mut pos)?;
                active_init = Some((tableidx, offset));
                skip_leb128(data, &mut pos); // reftype
                true
            }
            7 => {
                skip_leb128(data, &mut pos); // reftype
                true
            }
            _ => return Err("Invalid WASM: unsupported element segment mode".into()) };
        let (len, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut segment = Vec::with_capacity(len as usize);
        for _ in 0..len {
            if expr_items {
                segment.push(read_ref_const_expr(data, &mut pos)?);
            } else {
                let (func_idx, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                segment.push(Value::I32(func_idx as i32));
            }
        }
        if let Some((table_index, offset)) = active_init {
            active.push(ActiveElementSegment {
                table_index,
                offset,
                elem_index });
        }
        segments.push(segment);
    }
    Ok((segments, active))
}

#[allow(clippy::too_many_arguments)]
fn validate_code_bodies(
    code_sec: &[u8],
    func_type_indices: &[u32],
    types: &[(Vec<u8>, Vec<u8>)],
    func_sigs: &[(usize, usize)],
    global_count: usize,
    global_mutability: &[bool],
    data_count: usize,
    elem_count: usize,
    has_data_count_section: bool,
    uses_memory64: bool,
    uses_table64: bool,
) -> Result<(), String> {
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&code_sec[pos..]);
    pos += read;
    for func_idx in 0..count as usize {
        let (body_size, read) = read_leb128_u32(&code_sec[pos..]);
        pos += read;
        let body_end = pos
            .checked_add(body_size as usize)
            .ok_or_else(|| "Invalid WASM: code body size overflow".to_string())?;
        if body_end > code_sec.len() || body_end <= pos || code_sec[body_end - 1] != 0x0B {
            return Err("Invalid WASM: code body missing end opcode".into());
        }

        let (local_groups, read) = read_leb128_u32(&code_sec[pos..body_end]);
        pos += read;
        let mut local_count = 0usize;
        for _ in 0..local_groups {
            let (n, read) = read_leb128_u32(&code_sec[pos..body_end]);
            pos += read;
            if pos >= body_end {
                return Err("Invalid WASM: malformed local declaration".into());
            }
            pos += 1; // valtype
            local_count += n as usize;
        }

        let type_idx = func_type_indices
            .get(func_idx)
            .copied()
            .ok_or_else(|| "Invalid WASM: missing function type".to_string())?;
        let (param_count, result_arity) = types
            .get(type_idx as usize)
            .map(|(params, results)| (params.len(), results.len()))
            .ok_or_else(|| "Invalid WASM: function type index out of range".to_string())?;
        let total_locals = param_count + local_count;
        validate_instruction_stream(
            &code_sec[pos..body_end - 1],
            total_locals,
            func_sigs,
            types,
            result_arity,
            global_count,
            global_mutability,
            data_count,
            elem_count,
            has_data_count_section,
            uses_memory64,
            uses_table64,
        )?;
        pos = body_end;
    }
    if pos != code_sec.len() {
        return Err("Invalid WASM: trailing bytes after code bodies".into());
    }
    Ok(())
}

/// Arity-form WebAssembly validation (spec 3.4 + appendix "Validation
/// Algorithm"): a value-stack height plus control frames with
/// unreachable polymorphism. Value *types* are not checked — only
/// arities — which accepts every valid module (incl. rustc/clang
/// output) while still rejecting structurally broken bytecode.
struct CtrlFrame {
    start_height: usize,
    param_arity: usize,
    result_arity: usize,
    is_loop: bool,
    unreachable: bool }

struct ArityStack {
    height: usize,
    frames: Vec<CtrlFrame> }

impl ArityStack {
    fn new(result_arity: usize) -> Self {
        ArityStack {
            height: 0,
            frames: vec![CtrlFrame {
                start_height: 0,
                param_arity: 0,
                result_arity,
                is_loop: false,
                unreachable: false }] }
    }

    fn push(&mut self, n: usize) {
        self.height += n;
    }

    /// Pop `n` values. Inside an unreachable frame, popping below the
    /// frame base is polymorphic (always allowed) per the spec.
    fn pop(&mut self, n: usize, context: &str) -> Result<(), String> {
        let frame = self
            .frames
            .last()
            .ok_or_else(|| format!("Invalid WASM: no frame in {context}"))?;
        for _ in 0..n {
            if self.height > frame.start_height {
                self.height -= 1;
            } else if !frame.unreachable {
                return Err(format!("Invalid WASM: stack underflow in {context}"));
            }
        }
        Ok(())
    }

    fn set_unreachable(&mut self) {
        if let Some(frame) = self.frames.last_mut() {
            self.height = frame.start_height;
            frame.unreachable = true;
        }
    }

    fn push_frame(
        &mut self,
        param_arity: usize,
        result_arity: usize,
        is_loop: bool,
        context: &str,
    ) -> Result<(), String> {
        self.pop(param_arity, context)?;
        self.frames.push(CtrlFrame {
            start_height: self.height,
            param_arity,
            result_arity,
            is_loop,
            unreachable: false });
        self.push(param_arity);
        Ok(())
    }

    fn pop_frame(&mut self, context: &str) -> Result<CtrlFrame, String> {
        if self.frames.len() <= 1 {
            return Err(format!("Invalid WASM: unbalanced end in {context}"));
        }
        let (start, results, unreachable) = {
            let f = self.frames.last().unwrap();
            (f.start_height, f.result_arity, f.unreachable)
        };
        self.pop(results, context)?;
        if self.height != start && !unreachable {
            return Err(format!(
                "Invalid WASM: block leaves wrong stack height in {context}"
            ));
        }
        self.height = start;
        Ok(self.frames.pop().unwrap())
    }

    /// Branch-target arity: loops receive their params, blocks/ifs
    /// their results (spec: label types).
    fn label_arity(&self, depth: u32) -> Result<usize, String> {
        let idx = self
            .frames
            .len()
            .checked_sub(1 + depth as usize)
            .ok_or_else(|| "Invalid WASM: branch depth out of range".to_string())?;
        let f = &self.frames[idx];
        Ok(if f.is_loop {
            f.param_arity
        } else {
            f.result_arity
        })
    }
}

/// Signed LEB128, 33-bit (blocktype type indices).
fn read_sleb33(data: &[u8]) -> (i64, usize) {
    let mut result: i64 = 0;
    let mut shift = 0u32;
    let mut read = 0usize;
    loop {
        let Some(&byte) = data.get(read) else {
            return (result, read);
        };
        read += 1;
        result |= i64::from(byte & 0x7F) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 64 && byte & 0x40 != 0 {
                result |= -1i64 << shift;
            }
            return (result, read);
        }
        if shift >= 40 {
            return (result, read);
        }
    }
}

/// Blocktype: 0x40 = empty, a valtype shorthand = one result, or an
/// s33 type-section index carrying full (params, results).
fn decode_blocktype(
    code: &[u8],
    pos: &mut usize,
    types: &[(Vec<u8>, Vec<u8>)],
) -> Result<(usize, usize), String> {
    let b = *code
        .get(*pos)
        .ok_or_else(|| "Invalid WASM: truncated blocktype".to_string())?;
    if b == 0x40 {
        *pos += 1;
        return Ok((0, 0));
    }
    if (0x63..=0x7F).contains(&b) {
        *pos += 1;
        // (ref ht) / (ref null ht) shorthands carry a heaptype immediate,
        // which may be `(exact $x)` — two lebs, not one.
        if b == 0x63 || b == 0x64 {
            skip_heaptype(code, pos);
        }
        return Ok((0, 1));
    }
    let (idx, read) = read_sleb33(&code[*pos..]);
    *pos += read;
    let (params, results) = types
        .get(usize::try_from(idx).unwrap_or(usize::MAX))
        .ok_or_else(|| "Invalid WASM: blocktype type index out of range".to_string())?;
    Ok((params.len(), results.len()))
}

#[allow(clippy::too_many_arguments)]
fn validate_instruction_stream(
    code: &[u8],
    local_count: usize,
    func_sigs: &[(usize, usize)],
    types: &[(Vec<u8>, Vec<u8>)],
    result_arity: usize,
    global_count: usize,
    global_mutability: &[bool],
    data_count: usize,
    elem_count: usize,
    has_data_count_section: bool,
    uses_memory64: bool,
    uses_table64: bool,
) -> Result<(), String> {
    let mut pos = 0;
    let mut st = ArityStack::new(result_arity);
    // Pop a typed call's params and push its results.
    fn apply_sig(
        st: &mut ArityStack,
        sig: Option<&(usize, usize)>,
        context: &str,
        err: &str,
    ) -> Result<(), String> {
        let &(params, results) = sig.ok_or_else(|| err.to_string())?;
        st.pop(params, context)?;
        st.push(results);
        Ok(())
    }
    while pos < code.len() {
        let op = code[pos];
        pos += 1;
        match op {
            0x00 => st.set_unreachable(),
            0x01 => {}
            0x02 | 0x03 => {
                let (params, results) = decode_blocktype(code, &mut pos, types)?;
                st.push_frame(params, results, op == 0x03, "block")?;
            }
            0x04 => {
                st.pop(1, "if condition")?;
                let (params, results) = decode_blocktype(code, &mut pos, types)?;
                st.push_frame(params, results, false, "if")?;
            }
            0x05 => {
                let f = st.pop_frame("else")?;
                st.push_frame(f.param_arity, f.result_arity, false, "else")?;
            }
            0x08 => {
                // throw: tag arity isn't tracked here — the frame goes
                // unreachable either way.
                skip_leb128(code, &mut pos);
                st.set_unreachable();
            }
            // Legacy exception-handling block opcodes (try/catch/catch_all/
            // rethrow) — not supported; reject rather than silently mis-decode.
            0x06 | 0x07 | 0x09 | 0x19 => return Err(LEGACY_EH_UNSUPPORTED.into()),
            0x0A => {
                st.pop(1, "throw_ref")?;
                st.set_unreachable();
            }
            0x0B => {
                let f = st.pop_frame("end")?;
                st.push(f.result_arity);
            }
            0x0C | 0x0D => {
                let (depth, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                let arity = st.label_arity(depth)?;
                if op == 0x0D {
                    st.pop(1, "br_if condition")?;
                    st.pop(arity, "br_if")?;
                    st.push(arity);
                } else {
                    st.pop(arity, "br")?;
                    st.set_unreachable();
                }
            }
            0x0E => {
                st.pop(1, "br_table selector")?;
                let (count, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                for _ in 0..count {
                    let (depth, read) = read_leb128_u32(&code[pos..]);
                    pos += read;
                    st.label_arity(depth)?;
                }
                let (default_depth, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                let arity = st.label_arity(default_depth)?;
                st.pop(arity, "br_table")?;
                st.set_unreachable();
            }
            0x0F => {
                st.pop(result_arity, "return")?;
                st.set_unreachable();
            }
            0x10 | 0x12 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                apply_sig(
                    &mut st,
                    func_sigs.get(idx as usize),
                    "call",
                    "Invalid WASM: call function index out of range",
                )?;
                if op == 0x12 {
                    st.set_unreachable(); // return_call
                }
            }
            0x11 | 0x13 => {
                let (type_idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                skip_leb128(code, &mut pos); // table index
                st.pop(1, "call_indirect selector")?;
                let sig = types
                    .get(type_idx as usize)
                    .map(|(p, r)| (p.len(), r.len()));
                apply_sig(
                    &mut st,
                    sig.as_ref(),
                    "call_indirect",
                    "Invalid WASM: call_indirect type index out of range",
                )?;
                if op == 0x13 {
                    st.set_unreachable(); // return_call_indirect
                }
            }
            0x14 | 0x15 => {
                // call_ref / return_call_ref (function-references)
                let (type_idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                st.pop(1, "call_ref funcref")?;
                let sig = types
                    .get(type_idx as usize)
                    .map(|(p, r)| (p.len(), r.len()));
                apply_sig(
                    &mut st,
                    sig.as_ref(),
                    "call_ref",
                    "Invalid WASM: call_ref type index out of range",
                )?;
                if op == 0x15 {
                    st.set_unreachable();
                }
            }
            // Legacy `delegate` — not supported (see LEGACY_EH_UNSUPPORTED).
            0x18 => return Err(LEGACY_EH_UNSUPPORTED.into()),
            0x1A => st.pop(1, "drop")?,
            0x1B => {
                st.pop(3, "select")?;
                st.push(1);
            }
            0x1C => {
                let (count, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                pos = pos.saturating_add(count as usize).min(code.len());
                st.pop(3, "select_t")?;
                st.push(1);
            }
            0x1F => {
                // try_table: blocktype + catch clause vector
                let (params, results) = decode_blocktype(code, &mut pos, types)?;
                let (count, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                for _ in 0..count {
                    let kind = code.get(pos).copied().unwrap_or(0xFF);
                    pos += 1;
                    if kind == 0x00 || kind == 0x01 {
                        skip_leb128(code, &mut pos); // tag index
                    }
                    skip_leb128(code, &mut pos); // label index
                }
                st.push_frame(params, results, false, "try_table")?;
            }
            0x20 | 0x21 | 0x22 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= local_count {
                    return Err("Invalid WASM: local index out of range".into());
                }
                match op {
                    0x20 => st.push(1),
                    0x21 => st.pop(1, "local.set")?,
                    _ => {
                        st.pop(1, "local.tee")?;
                        st.push(1);
                    }
                }
            }
            0x23 | 0x24 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= global_count {
                    return Err("Invalid WASM: global index out of range".into());
                }
                if op == 0x24 {
                    if !global_mutability
                        .get(idx as usize)
                        .copied()
                        .unwrap_or(false)
                    {
                        return Err("Invalid WASM: global.set to immutable global".into());
                    }
                    st.pop(1, "global.set")?;
                } else {
                    st.push(1);
                }
            }
            0x25 => {
                skip_leb128(code, &mut pos);
                st.pop(1, "table.get")?;
                st.push(1);
            }
            0x26 => {
                skip_leb128(code, &mut pos);
                st.pop(2, "table.set")?;
            }
            0x28..=0x3E => {
                skip_memarg_or_memory_immediate(code, &mut pos, op);
                if matches!(op, 0x36..=0x3E) {
                    st.pop(2, "memory store")?;
                } else {
                    st.pop(1, "memory load")?;
                    st.push(1);
                }
            }
            0x3F => {
                skip_leb128(code, &mut pos);
                st.push(1); // memory.size
            }
            0x40 => {
                skip_leb128(code, &mut pos);
                st.pop(1, "memory.grow")?;
                st.push(1);
            }
            0x41 | 0x42 => {
                skip_leb128(code, &mut pos);
                st.push(1);
            }
            0x43 => {
                pos += 4;
                st.push(1);
            }
            0x44 => {
                pos += 8;
                st.push(1);
            }
            0x45..=0x66 => {
                let operands = if op == 0x45 || op == 0x50 { 1 } else { 2 };
                st.pop(operands, "comparison")?;
                st.push(1);
            }
            0x67..=0xA6 => {
                let operands = if matches!(op, 0x67..=0x69 | 0x79..=0x7B | 0x8B..=0x91 | 0x99..=0x9F)
                {
                    1
                } else {
                    2
                };
                st.pop(operands, "numeric operation")?;
                st.push(1);
            }
            0xA7..=0xC4 => {
                st.pop(1, "conversion")?;
                st.push(1);
            }
            0xD0 => {
                skip_leb128(code, &mut pos); // heaptype
                st.push(1);
            }
            0xD1 => {
                st.pop(1, "ref.is_null")?;
                st.push(1);
            }
            0xD2 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= func_sigs.len() {
                    return Err("Invalid WASM: ref.func index out of range".into());
                }
                st.push(1);
            }
            0xD3 => {
                st.pop(2, "ref.eq")?;
                st.push(1);
            }
            0xD4 => {
                st.pop(1, "ref.as_non_null")?;
                st.push(1);
            }
            // function-references: value-carrying null branches.
            //   `br_on_null $l     : [t* (ref null ht)] -> [t* (ref ht)]`
            //     iff `$l : [t*]`      — branches WITHOUT the ref (label takes t*)
            //   `br_on_non_null $l : [t* (ref null ht)] -> [t*]`
            //     iff `$l : [t* (ref ht)]` — branches WITH the ref
            0xD5 | 0xD6 => {
                let (depth, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                let arity = st.label_arity(depth)?;
                let name = if op == 0xD5 {
                    "br_on_null"
                } else {
                    "br_on_non_null"
                };
                // Both consume the nullable reference from the top of stack.
                st.pop(1, name)?;
                // The label's own arity must still be satisfied on the branch
                // edge; validate it is present, then restore for fallthrough.
                st.pop(arity, name)?;
                st.push(arity);
                if op == 0xD5 {
                    // Fallthrough re-types the ref as non-null and keeps it.
                    st.push(1);
                }
            }
            0xE0 => {
                skip_leb128(code, &mut pos); // continuation type index
                st.pop(1, "cont.new")?;
                st.push(1);
            }
            0xE1 => {
                skip_leb128(code, &mut pos); // source continuation type index
                skip_leb128(code, &mut pos); // destination continuation type index
                st.pop(1, "cont.bind")?;
                st.push(1);
            }
            0xE2 => {
                skip_leb128(code, &mut pos); // tag index
                st.pop(1, "suspend")?;
            }
            0xE3 => {
                skip_leb128(code, &mut pos); // continuation type index
                let _ = read_stack_switch_handlers(code, &mut pos);
                st.pop(2, "resume")?;
            }
            0xE4 => {
                skip_leb128(code, &mut pos); // continuation type index
                skip_leb128(code, &mut pos); // tag index
                let _ = read_stack_switch_handlers(code, &mut pos);
                st.pop(2, "resume_throw")?;
            }
            0xE5 => {
                // resume_throw_ref: cont type idx + resumetable. The exnref
                // is taken from the stack (no tag immediate).
                skip_leb128(code, &mut pos); // continuation type index
                let _ = read_stack_switch_handlers(code, &mut pos);
                st.pop(2, "resume_throw_ref")?;
            }
            0xE6 => {
                skip_leb128(code, &mut pos); // continuation type index
                skip_leb128(code, &mut pos); // tag index
                st.pop(2, "switch")?;
            }
            0xFB => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
                    // One typeidx: struct.new/new_default, array.new/new_default.
                    0x00..=0x01 | 0x06..=0x07 => {
                        skip_leb128(code, &mut pos);
                    }
                    // array.new_fixed is typeidx + N, not one immediate — the
                    // same two-leb shape the decoder reads.
                    0x08 => {
                        skip_leb128(code, &mut pos);
                        skip_leb128(code, &mut pos);
                    }
                    // typeidx + fieldidx / dataidx / elemidx.
                    0x02..=0x05 | 0x09..=0x0A | 0x12..=0x13 => {
                        skip_leb128(code, &mut pos);
                        skip_leb128(code, &mut pos);
                    }
                    // ref.test / ref.cast take a heaptype, which is two lebs
                    // when it is `(exact $x)`.
                    0x14..=0x17 => {
                        skip_heaptype(code, &mut pos);
                    }
                    // br_on_cast / br_on_cast_fail: castflags, labelidx, and
                    // TWO heaptypes — not a single immediate.
                    0x18..=0x19 => {
                        pos = pos.saturating_add(1).min(code.len()); // castflags
                        skip_leb128(code, &mut pos); // labelidx
                        skip_heaptype(code, &mut pos);
                        skip_heaptype(code, &mut pos);
                    }
                    0x1C => {
                        st.pop(1, "ref.i31")?;
                        st.push(1);
                    }
                    // Custom Descriptors: struct.new_desc, struct.new_default_desc
                    // and ref.get_desc each take a typeidx.
                    0x20..=0x22 => {
                        skip_leb128(code, &mut pos);
                    }
                    // ref.cast_desc_eq, in its non-null and nullable forms.
                    0x23..=0x24 => {
                        skip_heaptype(code, &mut pos);
                    }
                    // br_on_cast_desc_eq / _fail — same shape as br_on_cast.
                    0x25..=0x26 => {
                        pos = pos.saturating_add(1).min(code.len()); // castflags
                        skip_leb128(code, &mut pos); // labelidx
                        skip_heaptype(code, &mut pos);
                        skip_heaptype(code, &mut pos);
                    }
                    _ => {}
                }
            }
            0xFC => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
                    0x00..=0x07 => {
                        // non-trapping float-to-int conversions
                        st.pop(1, "trunc_sat")?;
                        st.push(1);
                    }
                    0x08 => {
                        if !has_data_count_section {
                            return Err("Invalid WASM: memory.init without data_count".into());
                        }
                        let (data_idx, read) = read_leb128_u32(&code[pos..]);
                        pos += read;
                        if data_idx as usize >= data_count {
                            return Err("Invalid WASM: memory.init data index out of range".into());
                        }
                        skip_leb128(code, &mut pos);
                        st.pop(3, "memory.init")?;
                    }
                    0x09 => {
                        let (data_idx, read) = read_leb128_u32(&code[pos..]);
                        pos += read;
                        if data_idx as usize >= data_count {
                            return Err("Invalid WASM: data.drop index out of range".into());
                        }
                    }
                    0x0A => {
                        skip_leb128(code, &mut pos);
                        skip_leb128(code, &mut pos);
                        st.pop(3, "memory.copy")?;
                    }
                    0x0B => {
                        skip_leb128(code, &mut pos);
                        st.pop(3, "memory.fill")?;
                    }
                    0x0C => {
                        let (elem_idx, read) = read_leb128_u32(&code[pos..]);
                        pos += read;
                        if elem_idx as usize >= elem_count {
                            return Err(
                                "Invalid WASM: table.init element index out of range".into()
                            );
                        }
                        skip_leb128(code, &mut pos);
                        st.pop(3, "table.init")?;
                    }
                    0x0D => {
                        skip_leb128(code, &mut pos); // elem.drop
                    }
                    0x0E => {
                        skip_leb128(code, &mut pos);
                        skip_leb128(code, &mut pos);
                        st.pop(3, "table.copy")?;
                    }
                    0x0F => {
                        skip_leb128(code, &mut pos);
                        st.pop(2, "table.grow")?;
                        st.push(1);
                    }
                    0x10 => {
                        skip_leb128(code, &mut pos);
                        st.push(1); // table.size
                    }
                    0x11 => {
                        skip_leb128(code, &mut pos);
                        st.pop(3, "table.fill")?;
                    }
                    _ => {}
                }
            }
            0xFD => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
                    0x00..=0x0B | 0x54..=0x5D => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        if sub == 0x0B || (0x58..=0x5B).contains(&sub) {
                            st.pop(2, "simd memory store")?;
                        } else {
                            st.pop(1, "simd memory load")?;
                            st.push(1);
                        }
                        if (0x54..=0x5B).contains(&sub) {
                            pos = pos.saturating_add(1).min(code.len());
                        }
                    }
                    0x0C => {
                        pos = pos.saturating_add(16).min(code.len());
                        st.push(1);
                    }
                    0x0D => {
                        pos = pos.saturating_add(16).min(code.len());
                        st.pop(2, "i8x16.shuffle")?;
                        st.push(1);
                    }
                    0x15..=0x22 => {
                        pos = pos.saturating_add(1).min(code.len());
                    }
                    _ => {}
                }
            }
            0xFE => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
                    0x00 => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        st.pop(2, "memory.atomic.notify")?;
                        st.push(1);
                    }
                    0x01 | 0x02 => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        st.pop(3, "memory.atomic.wait")?;
                        st.push(1);
                    }
                    0x03 => {
                        pos = pos.saturating_add(1).min(code.len());
                    }
                    0x10..=0x16 => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        st.pop(1, "atomic load")?;
                        st.push(1);
                    }
                    0x17..=0x1D => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        st.pop(2, "atomic store")?;
                    }
                    0x1E..=0x47 => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        st.pop(2, "atomic rmw")?;
                        st.push(1);
                    }
                    0x48..=0x4E => {
                        skip_memarg_for_memory_width(code, &mut pos, uses_memory64);
                        st.pop(3, "atomic cmpxchg")?;
                        st.push(1);
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }
    // The implicit end of the function body: exactly the declared
    // results remain (unless the tail is unreachable).
    if st.frames.len() != 1 {
        return Err("Invalid WASM: unclosed block at end of function body".into());
    }
    let body = &st.frames[0];
    if !body.unreachable && st.height != result_arity {
        return Err("Invalid WASM: function body leaves wrong stack height".into());
    }
    Ok(())
}

fn skip_memarg_or_memory_immediate(code: &[u8], pos: &mut usize, op: u8) {
    if op == 0x3F || op == 0x40 {
        skip_leb128(code, pos);
    } else {
        skip_memarg(code, pos);
    }
}

/// Decode a standard WASM module (e.g. from Rust/C compiler)
fn decode_standard_wasm(
    type_sec: &[u8],
    import_sec: &[u8],
    func_sec: &[u8],
    table_sec: &[u8],
    memory_sec: &[u8],
    export_sec: &[u8],
    elem_sec: &[u8],
    code_sec: &[u8],
    data_sec: &[u8],
    tag_sec: &[u8],
) -> Result<Vec<Chunk>, String> {
    // Parse type section to get function signatures
    let types = parse_type_section(type_sec);
    let func_type_indices = parse_function_section(func_sec);
    // Exception tags: each references a function type whose params are the
    // tag's payload; the arity lets `throw`/`catch` bind the right count.
    let tag_arities = parse_tag_section(tag_sec, &types);

    // Parse imports
    let imports = parse_import_section(import_sec);
    let import_func_count = imports.iter().filter(|(_, _, kind)| *kind == 0).count();

    // Parse exports to find function names
    let exports = parse_export_section(export_sec);
    let mut memory_min_pages = parse_imported_memory_min_pages(import_sec);
    memory_min_pages.extend(parse_memory_section(memory_sec));
    // Per-memory 64-bit index type, aligned with memory_min_pages (imported
    // memories first — treated as 32-bit — then declared section memories).
    let declared_is_64 = parse_memory_is_64(memory_sec);
    let imported_mem_count = memory_min_pages.len().saturating_sub(declared_is_64.len());
    let mut memory_is_64 = vec![false; imported_mem_count];
    memory_is_64.extend(declared_is_64);
    let mut memory_max_pages = parse_imported_memory_max_pages(import_sec);
    memory_max_pages.extend(parse_memory_max_section(memory_sec));
    let mut table_min_sizes = parse_imported_table_min_sizes(import_sec);
    table_min_sizes.extend(parse_table_section(table_sec));
    let declared_table_is_64 = parse_table_is_64(table_sec);
    let imported_table_count = table_min_sizes
        .len()
        .saturating_sub(declared_table_is_64.len());
    let mut table_is_64 = vec![false; imported_table_count];
    table_is_64.extend(declared_table_is_64);
    let (data_segments, active_data_segments) = parse_data_segments(data_sec)?;
    let (elem_segments, active_elem_segments) = parse_element_segments(elem_sec)?;
    let uses_memory64 = section_uses_memory64(memory_sec);
    let uses_table64 = section_uses_table64(table_sec);

    // Parse code section
    let mut cpos = 0;
    let (func_count, read) = read_leb128_u32(&code_sec[cpos..]);
    cpos += read;

    let mut chunks = Vec::new();

    // Create a script chunk that calls exported functions
    let mut script = Chunk::new("<script>");
    script.local_count = 0;

    for i in 0..func_count as usize {
        let (body_size, read) = read_leb128_u32(&code_sec[cpos..]);
        cpos += read;
        let _body_start = cpos;
        let body_end = cpos + body_size as usize;

        // Parse locals
        let (local_groups, read) = read_leb128_u32(&code_sec[cpos..]);
        cpos += read;
        let mut local_count: u32 = 0;
        for _ in 0..local_groups {
            let (count, read) = read_leb128_u32(&code_sec[cpos..]);
            cpos += read;
            cpos += 1; // type byte
            local_count += count;
        }

        // Get function name from exports
        let func_idx = import_func_count + i;
        let name = exports
            .iter()
            .find(|(_, idx)| *idx == func_idx)
            .map(|(n, _)| n.clone())
            .unwrap_or_else(|| format!("func_{}", i));

        // Get arity + result arity from the function's type signature.
        let type_idx = func_type_indices.get(i).copied().unwrap_or(func_idx as u32) as usize;
        let (arity, result_arity) = types
            .get(type_idx)
            .map(|(params, results)| (params.len() as u8, (results.len() as u8).max(1)))
            .unwrap_or((0, 1));

        // Translate WASM opcodes to our Chunk format
        let wasm_code = &code_sec[cpos..body_end.saturating_sub(1)]; // -1 for trailing 'end'
        let mut chunk = translate_wasm_to_chunk(
            wasm_code,
            &name,
            arity,
            local_count,
            import_func_count,
            uses_memory64,
            uses_table64,
            &types,
            &tag_arities,
        );
        chunk.result_arity = result_arity;
        // Imported functions have no receiver, so their WASM param count is
        // the arity (used by call_indirect's runtime type check).
        chunk.param_count = arity;
        chunk.memory_min_pages = memory_min_pages.clone();
        chunk.memory_max_pages = memory_max_pages.clone();
        chunk.memory_is_64 = memory_is_64.clone();
        chunk.table_min_sizes = table_min_sizes.clone();
        chunk.table_is_64 = table_is_64.clone();
        chunk.data_segments = data_segments.clone();
        chunk.elem_segments = elem_segments.clone();
        chunk.active_data_segments = active_data_segments.clone();
        chunk.active_elem_segments = active_elem_segments.clone();
        chunk.emit_op(Op::RETURN, 0);
        chunks.push(chunk);

        cpos = body_end;
    }

    // Add function imports to the script chunk. Memory/table/global imports
    // are represented in the decoded module metadata, not as callable host
    // functions.
    for (module, name, kind) in &imports {
        if *kind == 0 {
            script.add_import(module, name);
        }
    }
    script.memory_min_pages = memory_min_pages;
    script.memory_max_pages = memory_max_pages;
    script.memory_is_64 = memory_is_64;
    script.table_min_sizes = table_min_sizes;
    script.table_is_64 = table_is_64;
    script.data_segments = data_segments;
    script.elem_segments = elem_segments;
    script.active_data_segments = active_data_segments;
    script.active_elem_segments = active_elem_segments;

    // Insert script as chunk 0
    chunks.insert(0, script);

    Ok(chunks)
}

/// A pending catch clause of a `try_table` being decoded: the byte position of
/// its offset placeholder (patched with the handler position at the try's
/// `end`) and the spec catch label `L` it branches to.
struct EhClause {
    offset_pos: usize,
    label: u32 }

/// Per-source-block bookkeeping for the translate pass. `emitted_span` is how
/// many WASM blocks this one source block expands to in the emitted chunk: 1
/// for a plain block/loop/if, 2 for a `try_table` (which we wrap in an outer
/// `$skip` block so normal completion can branch past the catch trampolines).
/// The extra emitted level shifts `br` depths for branches that exit the try
/// region, so [`emitted_br_depth`] remaps every source `br`.
struct LabelInfo {
    emitted_span: u32,
    /// Present while decoding a `try_table` body; drives trampoline emission
    /// at the matching `end`.
    eh: Option<Vec<EhClause>> }

/// Translate a source `br N` (targets the label `n` levels out) into the
/// emitted branch depth, accounting for `$skip` wrappers inserted around any
/// `try_table` between here and the target. A source block's identity in the
/// emitted stream is its *primary* WASM block (the `try_table`/`block`/`loop`
/// itself); the `$skip` wrapper is an OUTER extra level. So branching to the
/// block `n` levels out crosses the full emitted span of the `n` innermost
/// blocks and lands on the target's primary block. When no try_tables are on
/// the stack every span is 1 and this is the identity.
fn emitted_br_depth(label_stack: &[LabelInfo], n: u32) -> u32 {
    let k = label_stack.len();
    let n = (n as usize).min(k.saturating_sub(1));
    label_stack[k - n..k].iter().map(|i| i.emitted_span).sum()
}

/// Translate WASM opcodes to our internal Chunk format.
/// Builds a proper constant pool and adjusts local indices.
fn translate_wasm_to_chunk(
    wasm: &[u8],
    name: &str,
    arity: u8,
    wasm_local_count: u32,
    _import_count: usize,
    uses_memory64: bool,
    uses_table64: bool,
    types: &[(Vec<u8>, Vec<u8>)],
    tag_arities: &[u8],
) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk.local_count = arity as u16 + wasm_local_count as u16;

    // Import the module's exception tags by a stable name so every function
    // chunk resolves the same tag index to the SAME load-time entity (a
    // `throw $t` in one function is caught by a `catch $t` in another). Build
    // a wasm-tag-index → chunk-tag-index map for `throw`/`try_table` decode.
    let tag_map: Vec<u16> = tag_arities
        .iter()
        .enumerate()
        .map(|(i, &ar)| chunk.import_exception_tag(format!("wasm:import:tag:{i}"), ar))
        .collect();

    let mut pos = 0;
    let mut label_stack: Vec<LabelInfo> = Vec::new();

    while pos < wasm.len() {
        let byte = wasm[pos];
        pos += 1;

        match byte {
            0x00 => chunk.emit_op(Op::HALT, 0),
            0x01 => {} // nop
            0x09 => {
                chunk.emit_op(Op::RETHROW, 0);
                read_emit_optional_memidx(&mut chunk, wasm, &mut pos);
            }
            // throw tagidx (0x08) — raise the tag with its payload on the stack.
            0x08 => {
                let (wasm_tag, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let chunk_tag = tag_map
                    .get(wasm_tag as usize)
                    .copied()
                    .unwrap_or(wasm_tag as u16);
                chunk.emit_op_u16(Op::THROW, chunk_tag, 0);
            }
            // throw_ref (0x0A) — re-raise the exnref on the stack.
            0x0A => chunk.emit_op(Op::THROW_REF, 0),
            // try_table blocktype catch* (0x1F) — WASM 3.0 structured EH.
            // Decode the up-front catch-clause vector, wrap the try_table in a
            // `$skip` block, and record each clause's spec target label; the
            // matching `end` emits the trampolines (see 0x0B).
            0x1F => {
                let result_count = read_block_result_count(wasm, &mut pos);
                let (clause_count, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let mut pairs: Vec<(u8, u16)> = Vec::with_capacity(clause_count as usize);
                let mut labels: Vec<u32> = Vec::with_capacity(clause_count as usize);
                for _ in 0..clause_count {
                    // kind: 0=catch 1=catch_ref 2=catch_all 3=catch_all_ref
                    // (identical to the VM CATCH_KIND_* values).
                    let kind = wasm.get(pos).copied().unwrap_or(2);
                    pos += 1;
                    let chunk_tag = if kind == 0x00 || kind == 0x01 {
                        let (wt, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        tag_map.get(wt as usize).copied().unwrap_or(wt as u16)
                    } else {
                        0
                    };
                    let (label, _) = read_leb128_u32(&wasm[pos..]);
                    skip_leb128(wasm, &mut pos);
                    pairs.push((kind, chunk_tag));
                    labels.push(label);
                }
                chunk.emit_block_typed(0, result_count); // $skip wrapper
                let offsets = chunk.emit_try_table_clauses(&pairs, 0);
                let clauses: Vec<EhClause> = offsets
                    .into_iter()
                    .zip(labels)
                    .map(|(offset_pos, label)| EhClause { offset_pos, label })
                    .collect();
                label_stack.push(LabelInfo {
                    emitted_span: 2,
                    eh: Some(clauses) });
            }

            // block blocktype — forward jump target
            0x02 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                chunk.emit_block_typed(0, result_count);
                label_stack.push(LabelInfo {
                    emitted_span: 1,
                    eh: None });
            }

            // loop blocktype — backward jump target
            0x03 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                chunk.emit_loop_typed(0, result_count);
                label_stack.push(LabelInfo {
                    emitted_span: 1,
                    eh: None });
            }

            // if blocktype — conditional block
            0x04 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                if result_count == 0 {
                    chunk.emit_if(0);
                } else {
                    chunk.emit_if_value(0);
                }
                label_stack.push(LabelInfo {
                    emitted_span: 1,
                    eh: None });
            }

            // else
            0x05 => {
                chunk.emit_else(0);
            }

            // end
            0x0B => {
                match label_stack.pop() {
                    // A `try_table` body's `end`: close the VM try_table, then
                    // build the catch trampolines. Normal completion branches
                    // past them to the enclosing `$skip` wrapper; on a caught
                    // exception the VM jumps to a trampoline `br L` that
                    // forwards to the spec target label (L=0 exits the region).
                    Some(LabelInfo {
                        eh: Some(clauses), ..
                    }) => {
                        chunk.emit_end(0); // close try_table (body done)
                        chunk.emit_br(0, 0); // normal path → $skip (now innermost)
                        for clause in &clauses {
                            // Patch this clause's forward offset to the handler.
                            let here = chunk.current_offset();
                            let jump = here as i32 - (clause.offset_pos as i32 + 2);
                            chunk.code[clause.offset_pos] = (jump >> 8) as u8;
                            chunk.code[clause.offset_pos + 1] = (jump & 0xff) as u8;
                            // Trampoline: spec `catch … L` → `br L`. The `$skip`
                            // wrapper (innermost here) makes L=0 exit the region;
                            // deeper labels add the wrapper level + remap.
                            let tramp = if clause.label == 0 {
                                0
                            } else {
                                1 + emitted_br_depth(&label_stack, clause.label - 1)
                            };
                            chunk.emit_br(tramp, 0);
                        }
                        chunk.emit_end(0); // close $skip
                    }
                    _ => chunk.emit_end(0) }
            }

            // br N — branch to Nth enclosing label (depth remapped for $skip)
            0x0C => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br(emitted_br_depth(&label_stack, depth), 0);
            }

            // br_if N — conditional branch
            0x0D => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br_if(emitted_br_depth(&label_stack, depth), 0);
            }
            0x0E => {
                // br_table — branch table
                let (count, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let mut depths = Vec::with_capacity(count as usize);
                for _ in 0..count {
                    let (depth, _) = read_leb128_u32(&wasm[pos..]);
                    skip_leb128(wasm, &mut pos);
                    depths.push(emitted_br_depth(&label_stack, depth));
                }
                let (default_depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br_table(&depths, emitted_br_depth(&label_stack, default_depth), 0);
            }
            0x0F => chunk.emit_op(Op::RETURN, 0),
            0x18 => {
                chunk.emit_op(Op::DELEGATE, 0);
                read_emit_optional_memidx(&mut chunk, wasm, &mut pos);
            }
            0x1A => chunk.emit_op(Op::DROP, 0),
            0x1B => chunk.emit_op(Op::SELECT, 0),

            // call funcidx — WASM direct call by function index
            0x10 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_call(idx as u16, 0, 0);
            }

            // local.get — slot 0 is the first argument, matching the VM.
            0x20 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_GET, idx as u16, 0);
            }
            // local.set
            0x21 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_SET, idx as u16, 0);
            }
            // local.tee — sets the local and LEAVES the value on the
            // stack (spec §4.4.5); mapping it to LOCAL_SET would drain
            // one stack slot per execution.
            0x22 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_TEE, idx as u16, 0);
            }

            // i32.const
            0x41 => {
                let (val, read) = read_leb128_i32(&wasm[pos..]);
                pos += read;
                chunk.emit_i32_const(val, 0);
            }

            // i64.const
            0x42 => {
                let (val, read) = read_leb128_i64(&wasm[pos..]);
                pos += read;
                chunk.emit_i64_const(val, 0);
            }

            // f64.const
            0x44 => {
                if pos + 8 <= wasm.len() {
                    let val = f64::from_le_bytes([
                        wasm[pos],
                        wasm[pos + 1],
                        wasm[pos + 2],
                        wasm[pos + 3],
                        wasm[pos + 4],
                        wasm[pos + 5],
                        wasm[pos + 6],
                        wasm[pos + 7],
                    ]);
                    pos += 8;
                    chunk.emit_f64_const(val, 0);
                }
            }

            // f32.const
            0x43 => {
                if pos + 4 <= wasm.len() {
                    let val = f32::from_le_bytes([
                        wasm[pos],
                        wasm[pos + 1],
                        wasm[pos + 2],
                        wasm[pos + 3],
                    ]);
                    pos += 4;
                    chunk.emit_f32_const(val, 0);
                }
            }

            // i32 arithmetic — ALL opcodes
            0x67 => chunk.emit_op(Op::I32_CLZ, 0),
            0x68 => chunk.emit_op(Op::I32_CTZ, 0),
            0x69 => chunk.emit_op(Op::I32_POPCNT, 0),
            0x6A => chunk.emit_op(Op::I32_ADD, 0),
            0x6B => chunk.emit_op(Op::I32_SUB, 0),
            0x6C => chunk.emit_op(Op::I32_MUL, 0),
            0x6D => chunk.emit_op(Op::I32_DIV_S, 0),
            0x6E => chunk.emit_op(Op::I32_DIV_U, 0),
            0x6F => chunk.emit_op(Op::I32_REM_S, 0),
            0x70 => chunk.emit_op(Op::I32_REM_U, 0),
            0x71 => chunk.emit_op(Op::I32_AND, 0),
            0x72 => chunk.emit_op(Op::I32_OR, 0),
            0x73 => chunk.emit_op(Op::I32_XOR, 0),
            0x74 => chunk.emit_op(Op::I32_SHL, 0),
            0x75 => chunk.emit_op(Op::I32_SHR_S, 0),
            0x76 => chunk.emit_op(Op::I32_SHR_U, 0),
            0x77 => chunk.emit_op(Op::I32_ROTL, 0),
            0x78 => chunk.emit_op(Op::I32_ROTR, 0),

            // i64 arithmetic — ALL opcodes
            0x79 => chunk.emit_op(Op::I64_CLZ, 0),
            0x7A => chunk.emit_op(Op::I64_CTZ, 0),
            0x7B => chunk.emit_op(Op::I64_POPCNT, 0),
            0x7C => chunk.emit_op(Op::I64_ADD, 0),
            0x7D => chunk.emit_op(Op::I64_SUB, 0),
            0x7E => chunk.emit_op(Op::I64_MUL, 0),
            0x7F => chunk.emit_op(Op::I64_DIV_S, 0),
            0x80 => chunk.emit_op(Op::I64_DIV_U, 0),
            0x81 => chunk.emit_op(Op::I64_REM_S, 0),
            0x82 => chunk.emit_op(Op::I64_REM_U, 0),
            0x83 => chunk.emit_op(Op::I64_AND, 0),
            0x84 => chunk.emit_op(Op::I64_OR, 0),
            0x85 => chunk.emit_op(Op::I64_XOR, 0),
            0x86 => chunk.emit_op(Op::I64_SHL, 0),
            0x87 => chunk.emit_op(Op::I64_SHR_S, 0),
            0x88 => chunk.emit_op(Op::I64_SHR_U, 0),
            0x89 => chunk.emit_op(Op::I64_ROTL, 0),
            0x8A => chunk.emit_op(Op::I64_ROTR, 0),

            // i64 comparison
            0x50 => chunk.emit_op(Op::I64_EQZ, 0),
            0x51 => chunk.emit_op(Op::I64_EQ, 0),
            0x52 => chunk.emit_op(Op::I64_NE, 0),
            0x53 => chunk.emit_op(Op::I64_LT_S, 0),
            0x54 => chunk.emit_op(Op::I64_LT_U, 0),
            0x55 => chunk.emit_op(Op::I64_GT_S, 0),
            0x56 => chunk.emit_op(Op::I64_GT_U, 0),
            0x57 => chunk.emit_op(Op::I64_LE_S, 0),
            0x58 => chunk.emit_op(Op::I64_LE_U, 0),
            0x59 => chunk.emit_op(Op::I64_GE_S, 0),
            0x5A => chunk.emit_op(Op::I64_GE_U, 0),

            // i32 comparison
            0x45 => chunk.emit_op(Op::I32_EQZ, 0),
            0x46 => chunk.emit_op(Op::I32_EQ, 0),
            0x47 => chunk.emit_op(Op::I32_NE, 0),
            0x48 => chunk.emit_op(Op::I32_LT_S, 0),
            0x49 => chunk.emit_op(Op::I32_LT_U, 0),
            0x4A => chunk.emit_op(Op::I32_GT_S, 0),
            0x4B => chunk.emit_op(Op::I32_GT_U, 0),
            0x4C => chunk.emit_op(Op::I32_LE_S, 0),
            0x4D => chunk.emit_op(Op::I32_LE_U, 0),
            0x4E => chunk.emit_op(Op::I32_GE_S, 0),
            0x4F => chunk.emit_op(Op::I32_GE_U, 0),

            // f64 arithmetic — ALL opcodes
            0xA0 => chunk.emit_op(Op::F64_ADD, 0),
            0xA1 => chunk.emit_op(Op::F64_SUB, 0),
            0xA2 => chunk.emit_op(Op::F64_MUL, 0),
            0xA3 => chunk.emit_op(Op::F64_DIV, 0),
            0xA4 => chunk.emit_op(Op::F64_MIN, 0),
            0xA5 => chunk.emit_op(Op::F64_MAX, 0),
            0xA6 => chunk.emit_op(Op::F64_COPYSIGN, 0),

            // f32 comparison (mapped to f64 ops — Vybe uses f64 internally)
            0x5B => chunk.emit_op(Op::F64_EQ, 0),
            0x5C => chunk.emit_op(Op::F64_NE, 0),
            0x5D => chunk.emit_op(Op::F64_LT, 0),
            0x5E => chunk.emit_op(Op::F64_GT, 0),
            0x5F => chunk.emit_op(Op::F64_LE, 0),
            0x60 => chunk.emit_op(Op::F64_GE, 0),

            // f64 comparison
            0x61 => chunk.emit_op(Op::F64_EQ, 0),
            0x62 => chunk.emit_op(Op::F64_NE, 0),
            0x63 => chunk.emit_op(Op::F64_LT, 0),
            0x64 => chunk.emit_op(Op::F64_GT, 0),
            0x65 => chunk.emit_op(Op::F64_LE, 0),
            0x66 => chunk.emit_op(Op::F64_GE, 0),

            // Memory — ALL load/store opcodes
            0x28 => {
                chunk.emit_op(Op::I32_LOAD, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x29 => {
                chunk.emit_op(Op::I64_LOAD, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2A => {
                chunk.emit_op(Op::F32_LOAD, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2B => {
                chunk.emit_op(Op::F64_LOAD, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2C => {
                chunk.emit_op(Op::I32_LOAD8_S, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2D => {
                chunk.emit_op(Op::I32_LOAD8_U, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2E => {
                chunk.emit_op(Op::I32_LOAD16_S, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2F => {
                chunk.emit_op(Op::I32_LOAD16_U, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x30 => {
                chunk.emit_op(Op::I64_LOAD8_S, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x31 => {
                chunk.emit_op(Op::I64_LOAD8_U, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x32 => {
                chunk.emit_op(Op::I64_LOAD16_S, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x33 => {
                chunk.emit_op(Op::I64_LOAD16_U, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x34 => {
                chunk.emit_op(Op::I64_LOAD32_S, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x35 => {
                chunk.emit_op(Op::I64_LOAD32_U, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x36 => {
                chunk.emit_op(Op::I32_STORE, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x37 => {
                chunk.emit_op(Op::I64_STORE, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x38 => {
                chunk.emit_op(Op::F32_STORE, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x39 => {
                chunk.emit_op(Op::F64_STORE, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3A => {
                chunk.emit_op(Op::I32_STORE8, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3B => {
                chunk.emit_op(Op::I32_STORE16, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3C => {
                chunk.emit_op(Op::I64_STORE8, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3D => {
                chunk.emit_op(Op::I64_STORE16, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3E => {
                chunk.emit_op(Op::I64_STORE32, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3F => {
                chunk.emit_op(Op::MEMORY_SIZE, 0);
                read_emit_optional_memidx(&mut chunk, wasm, &mut pos);
            }
            0x40 => {
                chunk.emit_op(Op::MEMORY_GROW, 0);
                read_emit_optional_memidx(&mut chunk, wasm, &mut pos);
            }

            // f32 arithmetic — ALL opcodes
            0x8B => chunk.emit_op(Op::F32_ABS, 0),
            0x8C => chunk.emit_op(Op::F32_NEG, 0),
            0x8D => chunk.emit_op(Op::F32_CEIL, 0),
            0x8E => chunk.emit_op(Op::F32_FLOOR, 0),
            0x8F => chunk.emit_op(Op::F32_TRUNC, 0),
            0x90 => chunk.emit_op(Op::F32_NEAREST, 0),
            0x91 => chunk.emit_op(Op::F32_SQRT, 0),
            0x92 => chunk.emit_op(Op::F64_ADD, 0), // f32.add (promoted)
            0x93 => chunk.emit_op(Op::F64_SUB, 0), // f32.sub (promoted)
            0x94 => chunk.emit_op(Op::F64_MUL, 0), // f32.mul (promoted)
            0x95 => chunk.emit_op(Op::F64_DIV, 0), // f32.div (promoted)
            0x96 => chunk.emit_op(Op::F32_MIN, 0),
            0x97 => chunk.emit_op(Op::F32_MAX, 0),
            0x98 => chunk.emit_op(Op::F32_COPYSIGN, 0),

            // f64 extra ops — ALL opcodes
            0x99 => chunk.emit_op(Op::F64_ABS, 0),
            0x9A => chunk.emit_op(Op::F64_NEG, 0),
            0x9B => chunk.emit_op(Op::F64_CEIL, 0),
            0x9C => chunk.emit_op(Op::F64_FLOOR, 0),
            0x9D => chunk.emit_op(Op::F64_TRUNC, 0),
            0x9E => chunk.emit_op(Op::F64_NEAREST, 0),
            0x9F => chunk.emit_op(Op::F64_SQRT, 0),

            // Conversions (WASM spec §5.3-binary.instructions 0xA7–0xBF)
            0xA7 => chunk.emit_op(Op::I32_WRAP_I64, 0), // i32.wrap_i64
            0xA8 => chunk.emit_op(Op::I32_TRUNC_F32_S, 0), // i32.trunc_f32_s
            0xA9 => chunk.emit_op(Op::I32_TRUNC_F32_U, 0), // i32.trunc_f32_u
            0xAA => chunk.emit_op(Op::I32_FROM_F64, 0), // i32.trunc_f64_s
            0xAB => chunk.emit_op(Op::I32_TRUNC_F64_U, 0), // i32.trunc_f64_u
            0xAC => chunk.emit_op(Op::I64_EXTEND_I32_S, 0), // i64.extend_i32_s
            0xAD => chunk.emit_op(Op::I64_EXTEND_I32_U, 0), // i64.extend_i32_u
            0xAE => chunk.emit_op(Op::I64_TRUNC_F32_S, 0), // i64.trunc_f32_s
            0xAF => chunk.emit_op(Op::I64_TRUNC_F32_U, 0), // i64.trunc_f32_u
            0xB0 => chunk.emit_op(Op::I64_TRUNC_F64_S, 0), // i64.trunc_f64_s
            0xB1 => chunk.emit_op(Op::I64_TRUNC_F64_U, 0), // i64.trunc_f64_u
            0xB2 => chunk.emit_op(Op::F32_CONVERT_I32_S, 0), // f32.convert_i32_s
            0xB3 => chunk.emit_op(Op::F32_CONVERT_I32_U, 0), // f32.convert_i32_u
            0xB4 => chunk.emit_op(Op::F32_CONVERT_I64_S, 0), // f32.convert_i64_s
            0xB5 => chunk.emit_op(Op::F32_CONVERT_I64_U, 0), // f32.convert_i64_u
            0xB6 => chunk.emit_op(Op::F32_DEMOTE_F64, 0), // f32.demote_f64
            0xB7 => chunk.emit_op(Op::F64_FROM_I32, 0), // f64.convert_i32_s
            0xB8 => chunk.emit_op(Op::F64_CONVERT_I32_U, 0), // f64.convert_i32_u
            0xB9 => chunk.emit_op(Op::F64_CONVERT_I64_S, 0), // f64.convert_i64_s
            0xBA => chunk.emit_op(Op::F64_CONVERT_I64_U, 0), // f64.convert_i64_u
            0xBB => chunk.emit_op(Op::F64_PROMOTE_F32, 0), // f64.promote_f32
            0xBC => chunk.emit_op(Op::I32_REINTERPRET_F32, 0), // i32.reinterpret_f32
            0xBD => chunk.emit_op(Op::I64_REINTERPRET_F64, 0), // i64.reinterpret_f64
            0xBE => chunk.emit_op(Op::F32_REINTERPRET_I32, 0), // f32.reinterpret_i32
            0xBF => chunk.emit_op(Op::F64_REINTERPRET_I64, 0), // f64.reinterpret_i64

            // Sign extension
            0xC0 => chunk.emit_op(Op::I32_EXTEND8_S, 0),
            0xC1 => chunk.emit_op(Op::I32_EXTEND16_S, 0),
            0xC2 => chunk.emit_op(Op::I64_EXTEND8_S, 0),
            0xC3 => chunk.emit_op(Op::I64_EXTEND16_S, 0),
            0xC4 => chunk.emit_op(Op::I64_EXTEND32_S, 0),

            // Reference types. `ref.null <ht>` keeps its heaptype immediate
            // rather than being collapsed onto two opcodes: an externref/funcref
            // null is the lenient plain null, a GC-heap null traps on the GC
            // accessors, and the VM tells them apart by reading the immediate
            // (`heaptype::is_gc_heap`). Previously the byte was read and thrown
            // away, and the GC case needed a custom opcode to survive.
            0xD0 => {
                let ht = wasm.get(pos).copied().unwrap_or(0);
                skip_leb128(wasm, &mut pos); // heaptype
                chunk.emit_ref_null(ht, 0);
            }
            0xD1 => chunk.emit_op(Op::REF_IS_NULL, 0),

            // global.get/set — a DECODED module's globals, named by index.
            // Not routed through `primitives::globals`: that is the compiler's
            // funnel for a module's own global namespace, and this crate sits
            // BELOW the compiler. Decoding someone else's module is not the
            // same operation.
            0x23 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let ci = chunk.intern_string_constant(&format!("__wasm_global_{}", idx));
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
            0x24 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let ci = chunk.intern_string_constant(&format!("__wasm_global_{}", idx));
                chunk.emit_op_u16(Op::GLOBAL_SET, ci, 0);
            }

            0x25 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u8(Op::TABLE_GET, idx as u8, 0);
            }
            0x26 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u8(Op::TABLE_SET, idx as u8, 0);
            }

            // call_indirect
            0x11 => {
                let (type_idx, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                let (table_idx, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                // `call_indirect` carries THREE operand bytes
                // (`U8_U8_U8`): argc, tableidx, and the expected result count.
                // The VM compares the callee's `result_arity` against that
                // third byte to enforce the spec's runtime type check, so it
                // must be emitted — otherwise the VM reads the next
                // instruction's byte as the result arity and every
                // result-returning callee trips a bogus signature mismatch.
                let (argc, expected_results) = types
                    .get(type_idx as usize)
                    .map(|(params, results)| (params.len() as u8, results.len() as u8))
                    .unwrap_or((0, 0));
                chunk.emit_op_u8_u8(Op::CALL_INDIRECT, argc, table_idx as u8, 0);
                chunk.emit(expected_results, 0);
            }

            // 0xFC prefix — nontrapping-float-to-int (0x00–0x07) + bulk-memory/table ops
            0xFC => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                match sub {
                    0x00 => chunk.emit_op(Op::I32_TRUNC_SAT_F32_S, 0),
                    0x01 => chunk.emit_op(Op::I32_TRUNC_SAT_F32_U, 0),
                    0x02 => chunk.emit_op(Op::I32_TRUNC_SAT_F64_S, 0),
                    0x03 => chunk.emit_op(Op::I32_TRUNC_SAT_F64_U, 0),
                    0x04 => chunk.emit_op(Op::I64_TRUNC_SAT_F32_S, 0),
                    0x05 => chunk.emit_op(Op::I64_TRUNC_SAT_F32_U, 0),
                    0x06 => chunk.emit_op(Op::I64_TRUNC_SAT_F64_S, 0),
                    0x07 => chunk.emit_op(Op::I64_TRUNC_SAT_F64_U, 0),
                    0x08 => {
                        let (data_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::MEMORY_INIT, data_idx as u8, 0);
                        read_emit_optional_memidx(&mut chunk, wasm, &mut pos);
                    }
                    0x09 => {
                        let (data_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::DATA_DROP, data_idx as u8, 0);
                    }
                    0x0A => {
                        chunk.emit_op(Op::MEMORY_COPY, 0);
                        read_emit_optional_memidx(&mut chunk, wasm, &mut pos); // dst memory
                        read_emit_optional_memidx(&mut chunk, wasm, &mut pos); // src memory
                    }
                    0x0B => {
                        chunk.emit_op(Op::MEMORY_FILL, 0);
                        read_emit_optional_memidx(&mut chunk, wasm, &mut pos);
                    }
                    0x0C => {
                        let (elem_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        let (table_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op(Op::TABLE_INIT, 0);
                        chunk.emit(elem_idx as u8, 0);
                        chunk.emit(table_idx as u8, 0);
                    }
                    0x0D => {
                        let (elem_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::ELEM_DROP, elem_idx as u8, 0);
                    }
                    0x0E => {
                        let (dst_table, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        let (src_table, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op(Op::TABLE_COPY, 0);
                        chunk.emit(dst_table as u8, 0);
                        chunk.emit(src_table as u8, 0);
                    }
                    0x0F => {
                        let (table_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::TABLE_GROW, table_idx as u8, 0);
                    }
                    0x10 => {
                        let (table_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::TABLE_SIZE, table_idx as u8, 0);
                    }
                    0x11 => {
                        let (table_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::TABLE_FILL, table_idx as u8, 0);
                    }
                    _ => {}
                }
            }

            // GC proposal prefix.
            0xFB => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_gc_prefixed(&mut chunk, sub, wasm, &mut pos);
            }

            // Stack-switching proposal.
            0xE0 => {
                skip_leb128(wasm, &mut pos); // continuation type index
                chunk.emit_op(Op::CONT_NEW, 0);
            }
            0xE1 => {
                skip_leb128(wasm, &mut pos); // source continuation type index
                skip_leb128(wasm, &mut pos); // destination continuation type index
                chunk.emit_op_u8(Op::CONT_BIND, 0, 0);
            }
            0xE2 => {
                let (tag_idx, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                chunk.emit_op_u16(Op::SUSPEND, tag_idx as u16, 0);
            }
            0xE3 => {
                skip_leb128(wasm, &mut pos); // continuation type index
                let handlers = read_stack_switch_handlers(wasm, &mut pos);
                let op_start = chunk.code.len();
                chunk.emit_op_u16(Op::RESUME, 0, 0);
                if !handlers.is_empty() {
                    chunk.stack_switch_handlers.insert(op_start, handlers);
                }
            }
            0xE4 => {
                skip_leb128(wasm, &mut pos); // continuation type index
                let (tag_idx, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                let handlers = read_stack_switch_handlers(wasm, &mut pos);
                let op_start = chunk.code.len();
                chunk.emit_op_u16(Op::RESUME_THROW, tag_idx as u16, 0);
                if !handlers.is_empty() {
                    chunk.stack_switch_handlers.insert(op_start, handlers);
                }
            }
            0xE5 => {
                // resume_throw_ref: cont type idx + resumetable; exnref
                // comes from the stack, no tag immediate.
                skip_leb128(wasm, &mut pos); // continuation type index
                let handlers = read_stack_switch_handlers(wasm, &mut pos);
                let op_start = chunk.code.len();
                chunk.emit_op(Op::RESUME_THROW_REF, 0);
                if !handlers.is_empty() {
                    chunk.stack_switch_handlers.insert(op_start, handlers);
                }
            }
            0xE6 => {
                skip_leb128(wasm, &mut pos); // continuation type index
                let (tag_idx, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                chunk.emit_op_u16(Op::SWITCH, tag_idx as u16, 0);
            }

            // SIMD and relaxed-SIMD proposal prefix.
            0xFD => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_simd_prefixed(&mut chunk, sub, wasm, &mut pos, uses_memory64);
            }

            // Threads/atomics proposal prefix.
            0xFE => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_threads_prefixed(&mut chunk, sub, wasm, &mut pos, uses_memory64);
            }

            // Unknown — skip
            _ => {}
        }
    }

    chunk
}

fn emit_gc_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    let Some(op) = u8::try_from(sub)
        .ok()
        .and_then(|s| Op::decode((0xFB) as u16, (s as u16) as u16))
    else {
        return;
    };
    match op {
        // Spec `struct.new $t` carries only a typeidx; the field count comes
        // from the type. Our bytecode is `(typeidx, count)` where count is
        // used ONLY by the dynamic (typeidx 0) object-literal form, so a
        // foreign module's typeidx goes in the first slot and the count is 0.
        // It used to be written into the COUNT slot, where the VM read it as
        // a key/value pair count and popped that many values off the stack.
        _ if op == Op::STRUCT_NEW => {
            let (type_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_struct_new(type_idx as u16, 0, 0);
        }
        // `struct.get $t i` / `get_s` / `get_u` / `set $t i` carry TWO
        // immediates (MVP.md §Instructions), not one. Reading a single leb
        // left the fieldidx to be decoded as the next opcode.
        //
        // The read variants map to the INDEXED form (`o.fields[i]`), which is
        // where a foreign `struct.new $t` now puts its values. `struct.set`
        // has no indexed counterpart in this VM yet — it stays name-keyed, so
        // a foreign struct.set does NOT write indexed storage. Consuming the
        // right number of bytes is the part that matters here; the semantic
        // half lands with the indexed-set work.
        _ if op == Op::STRUCT_GET || op == Op::STRUCT_GET_S || op == Op::STRUCT_GET_U => {
            skip_leb128(wasm, pos); // typeidx
            let (field_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_struct_field_op(Op::STRUCT_GET_U, 1, field_idx as u16, 0);
        }
        _ if op == Op::STRUCT_SET => {
            skip_leb128(wasm, pos); // typeidx
            let (field_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            // Indexed form now exists — a foreign `struct.set $t i` writes
            // the same storage `struct.get $t i` reads. typeidx 1 is a
            // non-zero marker: the VM only uses it to select the indexed
            // path, the fieldidx is what addresses the slot.
            chunk.emit_struct_field_op(Op::STRUCT_SET, 1, field_idx as u16, 0);
        }
        // `array.len` takes NO immediate — reading one consumed the first
        // byte of the following instruction.
        _ if op == Op::ARRAY_LENGTH => {
            chunk.emit_op(Op::ARRAY_LENGTH, 0);
        }
        _ if op == Op::STRUCT_NEW_DEFAULT
            || op == Op::ARRAY_NEW
            || op == Op::ARRAY_NEW_DEFAULT
            || op == Op::ARRAY_GET
            || op == Op::ARRAY_GET_S
            || op == Op::ARRAY_GET_U
            || op == Op::ARRAY_SET
            || op == Op::ARRAY_FILL =>
        {
            let (idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            match op.operand_format() {
                vybe_runtime::opcode::OperandFormat::U16 => chunk.emit_op_u16(op, idx as u16, 0),
                _ => chunk.emit_op(op, 0) }
        }
        _ if op == Op::ARRAY_NEW_FIXED => {
            // `array.new_fixed $t N` carries BOTH immediates, and our bytecode
            // now does too — the type index used to be read and thrown away,
            // which is what left every fixed array unstamped and therefore
            // exempt from the spec's bounds traps.
            let (type_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            let (extra, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_array_new_fixed(type_idx as u16, extra as u16, 0);
        }
        _ if op == Op::ARRAY_NEW_DATA
            || op == Op::ARRAY_NEW_ELEM
            || op == Op::ARRAY_INIT_DATA
            || op == Op::ARRAY_INIT_ELEM =>
        {
            let (type_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            let (segment_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_op(op, 0);
            chunk.emit((type_idx >> 8) as u8, 0);
            chunk.emit((type_idx & 0xff) as u8, 0);
            chunk.emit((segment_idx >> 8) as u8, 0);
            chunk.emit((segment_idx & 0xff) as u8, 0);
        }
        _ if op == Op::ARRAY_COPY => {
            skip_leb128(wasm, pos);
            skip_leb128(wasm, pos);
            chunk.emit_op(op, 0);
        }
        _ if op == Op::REF_TEST
            || op == Op::REF_TEST_NULL
            || op == Op::REF_CAST
            || op == Op::REF_CAST_NULL =>
        {
            // The heaptype survives now. It used to be read and thrown away,
            // replaced with the string `"__wasm_heaptype"` — which made every
            // decoded cast a test against a type that does not exist.
            let ht = read_heaptype(wasm, pos);
            chunk.emit_ref_type_op(op, ht, 0);
        }
        _ if op == Op::BR_ON_CAST
            || op == Op::BR_ON_CAST_FAIL
            || op == Op::BR_ON_CAST_DESC_EQ
            || op == Op::BR_ON_CAST_DESC_EQ_FAIL =>
        {
            skip_leb128(wasm, pos); // flags
            let (depth, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            skip_heaptype(wasm, pos);
            skip_heaptype(wasm, pos);
            let idx = chunk.add_constant(Value::String(Arc::from("__wasm_heaptype")));
            chunk.emit_op_u16(op, idx, 0);
            chunk.emit(depth as u8, 0);
        }
        // ── Custom Descriptors ────────────────────────────────────────────
        // `struct.new_desc` / `struct.new_default_desc` / `ref.get_desc` each
        // carry a typeidx, and `ref.cast_desc_eq` carries a heaptype. These
        // used to fall through to the operand-less default arm, so the bytes
        // of the immediate were decoded as if they were the NEXT instruction —
        // desynchronising everything after the first descriptor op in a module.
        _ if op == Op::STRUCT_NEW_DESC
            || op == Op::STRUCT_NEW_DEFAULT_DESC
            || op == Op::REF_GET_DESC =>
        {
            let (type_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_op_u16(op, type_idx as u16, 0);
        }
        _ if op == Op::REF_CAST_DESC_EQ || op == Op::REF_CAST_DESC_EQ_NULL => {
            skip_heaptype(wasm, pos);
            let idx = chunk.add_constant(Value::String(Arc::from("__wasm_heaptype")));
            chunk.emit_op_u16(op, idx, 0);
        }
        _ => chunk.emit_op(op, 0) }
}

fn emit_simd_prefixed(
    chunk: &mut Chunk,
    sub: u32,
    wasm: &[u8],
    pos: &mut usize,
    uses_memory64: bool,
) {
    if (0x100..=0x113).contains(&sub) {
        if let Some(op) = Op::decode((0xFD) as u16, (sub as u16) as u16) {
            chunk.emit_op(op, 0);
        }
        return;
    }

    let Some(op) = u8::try_from(sub)
        .ok()
        .and_then(|s| Op::decode((0xFD) as u16, (s as u16) as u16))
    else {
        return;
    };
    match op {
        _ if matches!(sub, 0x00..=0x0B | 0x54..=0x5D) => {
            chunk.emit_op(op, 0);
            copy_simd_memarg(wasm, pos, chunk, uses_memory64);
            if matches!(sub, 0x54..=0x5B) {
                let lane = wasm.get(*pos).copied().unwrap_or(0);
                *pos += 1;
                chunk.emit(lane, 0);
            }
        }
        _ if op == Op::V128_CONST => {
            chunk.emit_op(op, 0);
            for _ in 0..16 {
                let b = wasm.get(*pos).copied().unwrap_or(0);
                *pos += 1;
                chunk.emit(b, 0);
            }
        }
        _ if op == Op::I8X16_SHUFFLE => {
            chunk.emit_op(op, 0);
            for _ in 0..16 {
                let b = wasm.get(*pos).copied().unwrap_or(0);
                *pos += 1;
                chunk.emit(b, 0);
            }
        }
        _ if op.operand_format() == vybe_runtime::opcode::OperandFormat::U8 => {
            let lane = wasm.get(*pos).copied().unwrap_or(0);
            *pos += 1;
            chunk.emit_op_u8(op, lane, 0);
        }
        _ => chunk.emit_op(op, 0) }
}

fn copy_simd_memarg(wasm: &[u8], pos: &mut usize, chunk: &mut Chunk, uses_memory64: bool) {
    let (align, read) = read_leb128_u32(&wasm[*pos..]);
    *pos += read;
    let marker = 0x80 | if uses_memory64 { 0x100 } else { 0 };
    chunk.emit_leb_u32(align | marker, 0);
    copy_leb128(wasm, pos, chunk);
    if align & 0x40 != 0 {
        copy_leb128(wasm, pos, chunk);
    }
}

fn emit_threads_prefixed(
    chunk: &mut Chunk,
    sub: u32,
    wasm: &[u8],
    pos: &mut usize,
    uses_memory64: bool,
) {
    let Some(op) = u8::try_from(sub)
        .ok()
        .and_then(|s| Op::decode((0xFE) as u16, (s as u16) as u16))
    else {
        return;
    };
    match op {
        _ if op == Op::ATOMIC_FENCE => {
            chunk.emit_op(op, 0);
            let immediate = wasm.get(*pos).copied().unwrap_or(0);
            *pos = (*pos).saturating_add(1).min(wasm.len());
            chunk.emit(immediate, 0);
        }
        _ if op == Op::THREAD_SPAWN || op == Op::THREAD_JOIN => {
            chunk.emit_op(op, 0);
        }
        _ => {
            chunk.emit_op(op, 0);
            if uses_memory64 {
                copy_memarg64(wasm, pos, chunk);
            } else {
                copy_memarg(wasm, pos, chunk);
            }
        }
    }
}

fn copy_memarg64(wasm: &[u8], pos: &mut usize, chunk: &mut Chunk) {
    let (align, read) = read_leb128_u32(&wasm[*pos..]);
    *pos += read;
    chunk.emit_leb_u32(align | 0x80, 0);
    copy_leb128(wasm, pos, chunk);
}

fn copy_memarg(wasm: &[u8], pos: &mut usize, chunk: &mut Chunk) {
    copy_leb128(wasm, pos, chunk);
    copy_leb128(wasm, pos, chunk);
}

fn copy_leb128(wasm: &[u8], pos: &mut usize, chunk: &mut Chunk) {
    while *pos < wasm.len() {
        let byte = wasm[*pos];
        *pos += 1;
        chunk.emit(byte, 0);
        if byte & 0x80 == 0 {
            break;
        }
    }
}

fn read_leb128_i64(data: &[u8]) -> (i64, usize) {
    let mut result = 0i64;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        result |= ((byte & 0x7F) as i64) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 64 && (byte & 0x40 != 0) {
                result |= !0i64 << shift;
            }
            break;
        }
    }
    (result, pos)
}

fn read_leb128_u64_local(data: &[u8]) -> (u64, usize) {
    let mut result = 0u64;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        result |= ((byte & 0x7f) as u64) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    (result, pos)
}

/// Parse the tag section (id 13). Each tag entry is `attribute(u8, 0x00 =
/// exception) + type_idx(leb)`; the tag's payload arity is the referenced
/// function type's parameter count. Returns arity per tag, in index order.
fn parse_tag_section(data: &[u8], types: &[(Vec<u8>, Vec<u8>)]) -> Vec<u8> {
    if data.is_empty() {
        return Vec::new();
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut arities = Vec::with_capacity(count as usize);
    for _ in 0..count {
        if pos >= data.len() {
            break;
        }
        pos += 1; // attribute byte (0x00)
        let (type_idx, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let arity = types
            .get(type_idx as usize)
            .map(|(params, _)| params.len() as u8)
            .unwrap_or(0);
        arities.push(arity);
    }
    arities
}

/// Parse the type section into one `(params, results)` entry per type index.
///
/// The section is `vec(rectype)`, but the TYPE INDEX SPACE counts each subtype
/// INSIDE a rec group — one `rec` of four subtypes occupies indices 0..3 while
/// being a single element of the vector. Struct and array types therefore get a
/// placeholder entry rather than being skipped: dropping them would shift every
/// later function type's index, and those indices are what `call` and
/// blocktypes resolve against.
///
/// This used to look only for `0x60` and, on anything else, advance a SINGLE
/// byte and continue — which neither skipped the type nor kept alignment, so a
/// module carrying GC types desynchronised the parse from its first struct on.
/// Encodings per `proposals/gc/proposals/gc/MVP.md` §Binary Format.
fn parse_type_section(data: &[u8]) -> Vec<(Vec<u8>, Vec<u8>)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut types = Vec::new();
    for _ in 0..count {
        if pos >= data.len() {
            break;
        }
        if data[pos] == GC_REC {
            pos += 1;
            let (group_len, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            for _ in 0..group_len {
                if !parse_subtype(data, &mut pos, &mut types) {
                    return types;
                }
            }
        } else if !parse_subtype(data, &mut pos, &mut types) {
            return types;
        }
    }
    types
}

/// One `subtype`: an optional `sub`/`sub final` header, then a composite type.
/// Returns false if the bytes ran out, so the caller stops rather than looping
/// on a truncated section.
fn parse_subtype(data: &[u8], pos: &mut usize, types: &mut Vec<(Vec<u8>, Vec<u8>)>) -> bool {
    let Some(&tag) = data.get(*pos) else {
        return false;
    };
    if tag == GC_SUB || tag == GC_SUB_FINAL {
        *pos += 1;
        // vec(typeidx) of supertypes.
        let (supers, read) = read_leb128_u32(&data[*pos..]);
        *pos += read;
        for _ in 0..supers {
            skip_leb128(data, pos);
        }
        // Custom Descriptors may insert `(descriptor $x)` / `(describes $x)`
        // between the supertype vector and the composite type.
        while matches!(data.get(*pos), Some(&CD_DESCRIPTOR) | Some(&CD_DESCRIBES)) {
            *pos += 1;
            skip_leb128(data, pos);
        }
    }
    parse_comptype(data, pos, types)
}

/// One `comptype`: `func`, `struct` or `array`.
fn parse_comptype(data: &[u8], pos: &mut usize, types: &mut Vec<(Vec<u8>, Vec<u8>)>) -> bool {
    let Some(&tag) = data.get(*pos) else {
        return false;
    };
    *pos += 1;
    match tag {
        TYPE_FUNC => {
            let (param_count, read) = read_leb128_u32(&data[*pos..]);
            *pos += read;
            let mut params = Vec::with_capacity(param_count as usize);
            for _ in 0..param_count {
                match read_value_type(data, pos) {
                    Some(byte) => params.push(byte),
                    None => return false }
            }
            let (result_count, read) = read_leb128_u32(&data[*pos..]);
            *pos += read;
            let mut results = Vec::with_capacity(result_count as usize);
            for _ in 0..result_count {
                match read_value_type(data, pos) {
                    Some(byte) => results.push(byte),
                    None => return false }
            }
            types.push((params, results));
            true
        }
        GC_STRUCT => {
            let (field_count, read) = read_leb128_u32(&data[*pos..]);
            *pos += read;
            for _ in 0..field_count {
                if !skip_field_type(data, pos) {
                    return false;
                }
            }
            // Occupies a type index without being callable.
            types.push((Vec::new(), Vec::new()));
            true
        }
        GC_ARRAY => {
            if !skip_field_type(data, pos) {
                return false;
            }
            types.push((Vec::new(), Vec::new()));
            true
        }
        // An unknown composite tag means the encoding moved on without us;
        // stopping beats returning a table whose indices are quietly wrong.
        _ => false }
}

/// `fieldtype = storagetype mutability`. Storage adds the packed types `i8`
/// and `i16`, which are legal ONLY here and never as a value type.
fn skip_field_type(data: &[u8], pos: &mut usize) -> bool {
    match data.get(*pos) {
        Some(&PACKED_I8) | Some(&PACKED_I16) => *pos += 1,
        Some(_) => {
            if read_value_type(data, pos).is_none() {
                return false;
            }
        }
        None => return false }
    // Mutability byte.
    if *pos >= data.len() {
        return false;
    }
    *pos += 1;
    true
}

/// One value type, returning its leading byte.
///
/// `(ref ht)` (0x64) and `(ref null ht)` (0x63) carry a heaptype immediate
/// encoded as `s33`; every other value type is a single byte.
fn read_value_type(data: &[u8], pos: &mut usize) -> Option<u8> {
    let &byte = data.get(*pos)?;
    *pos += 1;
    if byte == 0x63 || byte == 0x64 {
        // The heaptype may itself be `(exact $x)`, which is two lebs.
        skip_heaptype(data, pos);
    }
    Some(byte)
}

fn parse_memory_section(data: &[u8]) -> Vec<u64> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut memories = Vec::with_capacity(count as usize);
    for _ in 0..count {
        if pos >= data.len() {
            break;
        }
        let flags = data[pos];
        pos += 1;
        let is_memory64 = flags & 0x04 != 0;
        let has_max = flags & 0x01 != 0;
        let min = if is_memory64 {
            let (value, read) = read_leb128_u64_local(&data[pos..]);
            pos += read;
            value
        } else {
            let (value, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            value as u64
        };
        if has_max {
            let (_, read) = if is_memory64 {
                read_leb128_u64_local(&data[pos..])
            } else {
                let (value, read) = read_leb128_u32(&data[pos..]);
                (value as u64, read)
            };
            pos += read;
        }
        memories.push(min);
    }
    memories
}

fn read_limits(data: &[u8], pos: &mut usize) -> (u64, Option<u64>) {
    if *pos >= data.len() {
        return (0, None);
    }
    let flags = data[*pos];
    *pos += 1;
    let is_64 = flags & 0x04 != 0;
    let has_max = flags & 0x01 != 0;
    let min = if is_64 {
        let (value, read) = read_leb128_u64_local(&data[*pos..]);
        *pos += read;
        value
    } else {
        let (value, read) = read_leb128_u32(&data[*pos..]);
        *pos += read;
        value as u64
    };
    let max = if has_max {
        let (value, read) = if is_64 {
            read_leb128_u64_local(&data[*pos..])
        } else {
            let (value, read) = read_leb128_u32(&data[*pos..]);
            (value as u64, read)
        };
        *pos += read;
        Some(value)
    } else {
        None
    };
    (min, max)
}

fn read_limits_min(data: &[u8], pos: &mut usize) -> u64 {
    let (min, _) = read_limits(data, pos);
    min
}

fn parse_memory_max_section(data: &[u8]) -> Vec<Option<u64>> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut memories = Vec::with_capacity(count as usize);
    for _ in 0..count {
        let (_, max) = read_limits(data, &mut pos);
        memories.push(max);
    }
    memories
}

fn parse_imported_memory_min_pages(data: &[u8]) -> Vec<u64> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut memories = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + nlen as usize;
        if pos >= data.len() {
            break;
        }
        let kind = data[pos];
        pos += 1;
        match kind {
            0 => skip_leb128(data, &mut pos),
            1 => {
                skip_leb128(data, &mut pos); // reftype
                let _ = read_limits_min(data, &mut pos);
            }
            2 => memories.push(read_limits_min(data, &mut pos)),
            3 => {
                skip_leb128(data, &mut pos); // valtype
                pos = pos.saturating_add(1).min(data.len()); // mutability
            }
            _ => break }
    }
    memories
}

fn parse_imported_memory_max_pages(data: &[u8]) -> Vec<Option<u64>> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut memories = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + nlen as usize;
        if pos >= data.len() {
            break;
        }
        let kind = data[pos];
        pos += 1;
        match kind {
            0 => skip_leb128(data, &mut pos),
            1 => {
                skip_leb128(data, &mut pos);
                let _ = read_limits(data, &mut pos);
            }
            2 => {
                let (_, max) = read_limits(data, &mut pos);
                memories.push(max);
            }
            3 => {
                skip_leb128(data, &mut pos);
                pos = pos.saturating_add(1).min(data.len());
            }
            _ => break }
    }
    memories
}

fn parse_imported_table_min_sizes(data: &[u8]) -> Vec<u64> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut tables = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read + nlen as usize;
        if pos >= data.len() {
            break;
        }
        let kind = data[pos];
        pos += 1;
        match kind {
            0 => skip_leb128(data, &mut pos),
            1 => {
                skip_leb128(data, &mut pos); // reftype
                tables.push(read_limits_min(data, &mut pos));
            }
            2 => {
                let _ = read_limits_min(data, &mut pos);
            }
            3 => {
                skip_leb128(data, &mut pos); // valtype
                pos = pos.saturating_add(1).min(data.len()); // mutability
            }
            _ => break }
    }
    tables
}

fn parse_table_section(data: &[u8]) -> Vec<u64> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut tables = Vec::with_capacity(count as usize);
    for _ in 0..count {
        if pos + 1 >= data.len() {
            break;
        }
        pos += 1; // reftype
        let flags = data[pos];
        pos += 1;
        let is_table64 = flags & 0x04 != 0;
        let has_max = flags & 0x01 != 0;
        let min = if is_table64 {
            let (value, read) = read_leb128_u64_local(&data[pos..]);
            pos += read;
            value
        } else {
            let (value, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            value as u64
        };
        if has_max {
            let (_, read) = if is_table64 {
                read_leb128_u64_local(&data[pos..])
            } else {
                let (value, read) = read_leb128_u32(&data[pos..]);
                (value as u64, read)
            };
            pos += read;
        }
        tables.push(min);
    }
    tables
}

/// Per-memory index type from the memory section: `true` = 64-bit. Aligns
/// with `parse_memory_section`. The VM reads this (via `chunk.memory_is_64`)
/// to decide load/store address width — memory64 adds no new opcodes.
fn parse_memory_is_64(data: &[u8]) -> Vec<bool> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut out = Vec::with_capacity(count as usize);
    for _ in 0..count {
        if pos >= data.len() {
            break;
        }
        let flags = data[pos];
        pos += 1;
        let is_memory64 = flags & 0x04 != 0;
        let has_max = flags & 0x01 != 0;
        out.push(is_memory64);
        // Skip min (and max if present), width per index type, to reach the
        // next entry.
        for _ in 0..(1 + usize::from(has_max)) {
            if is_memory64 {
                let (_, r) = read_leb128_u64_local(&data[pos..]);
                pos += r;
            } else {
                let (_, r) = read_leb128_u32(&data[pos..]);
                pos += r;
            }
        }
    }
    out
}

fn section_uses_memory64(data: &[u8]) -> bool {
    if data.is_empty() {
        return false;
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    for _ in 0..count {
        if pos >= data.len() {
            break;
        }
        let flags = data[pos];
        pos += 1;
        let is_memory64 = flags & 0x04 != 0;
        let has_max = flags & 0x01 != 0;
        if is_memory64 {
            return true;
        }
        skip_leb128(data, &mut pos);
        if has_max {
            skip_leb128(data, &mut pos);
        }
    }
    false
}

/// Per-table index type from the table section: `true` = 64-bit (table64).
/// Aligns with `parse_table_section`. table64 adds no new opcodes.
fn parse_table_is_64(data: &[u8]) -> Vec<bool> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut out = Vec::with_capacity(count as usize);
    for _ in 0..count {
        if pos + 1 >= data.len() {
            break;
        }
        pos += 1; // element type
        let flags = data[pos];
        pos += 1;
        out.push(flags & 0x04 != 0);
        let has_max = flags & 0x01 != 0;
        skip_leb128(data, &mut pos);
        if has_max {
            skip_leb128(data, &mut pos);
        }
    }
    out
}

fn section_uses_table64(data: &[u8]) -> bool {
    if data.is_empty() {
        return false;
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    for _ in 0..count {
        if pos + 1 >= data.len() {
            break;
        }
        pos += 1; // element type
        let flags = data[pos];
        pos += 1;
        let is_table64 = flags & 0x04 != 0;
        let has_max = flags & 0x01 != 0;
        if is_table64 {
            return true;
        }
        skip_leb128(data, &mut pos);
        if has_max {
            skip_leb128(data, &mut pos);
        }
    }
    false
}

fn parse_function_section(data: &[u8]) -> Vec<u32> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut funcs = Vec::with_capacity(count as usize);
    for _ in 0..count {
        let (type_idx, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        funcs.push(type_idx);
    }
    funcs
}

fn parse_import_section(data: &[u8]) -> Vec<(String, String, u8)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut imports = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let module = std::str::from_utf8(&data[pos..pos + mlen as usize])
            .unwrap_or("")
            .to_string();
        pos += mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize])
            .unwrap_or("")
            .to_string();
        pos += nlen as usize;
        let kind = normalize_import_kind(data[pos]);
        pos += 1;
        skip_import_descriptor(data, &mut pos, kind);
        imports.push((module, name, kind));
    }
    imports
}

fn parse_export_section(data: &[u8]) -> Vec<(String, usize)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut exports = Vec::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize])
            .unwrap_or("")
            .to_string();
        pos += nlen as usize;
        let kind = data[pos];
        pos += 1;
        let (idx, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        if kind == 0 {
            // function export
            exports.push((name, idx as usize));
        }
    }
    exports
}

fn skip_leb128(data: &[u8], pos: &mut usize) {
    while *pos < data.len() {
        let byte = data[*pos];
        *pos += 1;
        if byte & 0x80 == 0 {
            break;
        }
    }
}

fn read_stack_switch_handlers(data: &[u8], pos: &mut usize) -> Vec<StackSwitchHandler> {
    let (handler_count, read) = read_leb128_u32(&data[*pos..]);
    *pos += read;
    let mut handlers = Vec::with_capacity(handler_count as usize);
    for _ in 0..handler_count {
        let kind = data.get(*pos).copied().unwrap_or(0);
        *pos = (*pos).saturating_add(1).min(data.len());
        let (tag_index, read) = read_leb128_u32(&data[*pos..]);
        *pos += read;
        let label_index = if kind == 0 {
            let (label, read) = read_leb128_u32(&data[*pos..]);
            *pos += read;
            label
        } else {
            0
        };
        handlers.push(StackSwitchHandler {
            kind,
            tag_index,
            label_index });
    }
    handlers
}

#[allow(dead_code)]
fn skip_const_expr(data: &[u8], pos: &mut usize) -> Result<(), String> {
    while *pos < data.len() {
        let op = data[*pos];
        *pos += 1;
        match op {
            0x0B => return Ok(()),
            0x41 | 0x42 | 0x23 | 0xD2 => skip_leb128(data, pos),
            0x43 => *pos += 4,
            0x44 => *pos += 8,
            0xD0 => *pos += 1,
            0xFC => skip_leb128(data, pos),
            _ => {
                return Err(format!(
                    "Invalid WASM: unsupported const expr opcode 0x{op:02x}"
                ));
            }
        }
        if *pos > data.len() {
            return Err("Invalid WASM: truncated const expression".into());
        }
    }
    Err("Invalid WASM: unterminated const expression".into())
}

fn read_i32_const_expr_as_u64(data: &[u8], pos: &mut usize) -> Result<u64, String> {
    if data.get(*pos).copied() != Some(0x41) {
        return Err("Invalid WASM: active segment offset must be i32.const".into());
    }
    *pos += 1;
    let (value, read) = read_leb128_i32(&data[*pos..]);
    *pos += read;
    if value < 0 {
        return Err("Invalid WASM: active segment offset must be non-negative".into());
    }
    if data.get(*pos).copied() != Some(0x0B) {
        return Err("Invalid WASM: active segment offset expression missing end".into());
    }
    *pos += 1;
    Ok(value as u64)
}

fn read_ref_const_expr(data: &[u8], pos: &mut usize) -> Result<Value, String> {
    if *pos >= data.len() {
        return Err("Invalid WASM: truncated element expression".into());
    }
    let op = data[*pos];
    *pos += 1;
    let value = match op {
        0xD0 => {
            if *pos >= data.len() {
                return Err("Invalid WASM: truncated ref.null expression".into());
            }
            *pos += 1;
            Value::Null
        }
        0xD2 => {
            let (func_idx, read) = read_leb128_u32(&data[*pos..]);
            *pos += read;
            Value::I32(func_idx as i32)
        }
        _ => {
            return Err(format!(
                "Invalid WASM: unsupported element expression opcode 0x{op:02x}"
            ));
        }
    };
    if data.get(*pos).copied() != Some(0x0B) {
        return Err("Invalid WASM: element expression missing end".into());
    }
    *pos += 1;
    Ok(value)
}

fn read_emit_leb_u32(chunk: &mut Chunk, data: &[u8], pos: &mut usize) -> u32 {
    let (value, read) = read_leb128_u32(&data[*pos..]);
    for byte in &data[*pos..*pos + read] {
        chunk.emit(*byte, 0);
    }
    *pos += read;
    value
}

fn read_emit_optional_memidx(chunk: &mut Chunk, data: &[u8], pos: &mut usize) -> u32 {
    let (value, read) = read_leb128_u32(&data[*pos..]);
    *pos += read;
    emit_explicit_memidx(chunk, value);
    value
}

/// Emit the VM's multi-memory selector.
///
/// SINGLE SOURCE OF TRUTH for the layout is the VM
/// (`dispatch::read_optional_memidx_immediate`): a **fixed 4-byte** block
/// `0xEE 0x00 <memidx u16 BE>`. VM instructions are always 4 bytes, so the
/// selector must be 4 too or the following instruction loses alignment and
/// execution desyncs. A LEB-encoded memidx here would be 3 bytes for small
/// values — one short — which ran the interpreter off the end of the code.
fn emit_explicit_memidx(chunk: &mut Chunk, value: u32) {
    chunk.emit(0xEE, 0);
    chunk.emit(0x00, 0);
    chunk.emit((value >> 8) as u8, 0);
    chunk.emit((value & 0xFF) as u8, 0);
}

fn read_emit_memarg(chunk: &mut Chunk, data: &[u8], pos: &mut usize) {
    let align = read_emit_leb_u32(chunk, data, pos);
    let _offset = read_emit_leb_u32(chunk, data, pos);
    if align & 0x40 != 0 {
        let _memidx = read_emit_leb_u32(chunk, data, pos);
    }
}

fn read_emit_memarg_for_memory_width(
    chunk: &mut Chunk,
    data: &[u8],
    pos: &mut usize,
    is_memory64: bool,
) {
    if is_memory64 {
        read_emit_memarg64(chunk, data, pos);
    } else {
        read_emit_memarg(chunk, data, pos);
    }
}

fn read_emit_memarg64(chunk: &mut Chunk, data: &[u8], pos: &mut usize) {
    let align = read_emit_leb_u32(chunk, data, pos);
    let (_offset, read) = read_leb128_u64_local(&data[*pos..]);
    for byte in &data[*pos..*pos + read] {
        chunk.emit(*byte, 0);
    }
    *pos += read;
    if align & 0x40 != 0 {
        let _memidx = read_emit_leb_u32(chunk, data, pos);
    }
}

#[allow(dead_code)]
fn simd_memory_opcode_name(sub: u32) -> &'static str {
    match sub {
        0x00 => "v128.load",
        0x0B => "v128.store",
        _ => "SIMD memory operation" }
}

#[allow(dead_code)]
fn atomic_opcode_name(sub: u32) -> &'static str {
    match sub {
        0x10 => "i32.atomic.load",
        0x17 => "i32.atomic.store",
        _ => "atomic memory operation" }
}

fn skip_memarg(data: &[u8], pos: &mut usize) {
    skip_leb128(data, pos); // align
    skip_leb128(data, pos); // offset
}

fn skip_memarg_for_memory_width(data: &[u8], pos: &mut usize, _is_memory64: bool) {
    skip_memarg(data, pos);
}

/// One heaptype immediate.
///
/// Normally an `s33` — negative values are the abstract heap types, positive
/// ones are type indices. Custom Descriptors adds `0x62 x:u32` for `(exact
/// $x)`, so a leading `HEAPTYPE_EXACT` is followed by a SECOND leb that must
/// also be consumed. Missing it leaves the type index to be decoded as an
/// instruction.
/// Decode a heaptype immediate: a signed LEB where negative is one of the
/// abstract types and non-negative is a type index. `(exact $t)` — Custom
/// Descriptors — narrows to the same index; exactness is a property of the
/// cast, not of the type it names.
fn read_heaptype(data: &[u8], pos: &mut usize) -> vybe_runtime::opcode::heaptype::HeapType {
    if data.get(*pos) == Some(&HEAPTYPE_EXACT) {
        *pos += 1;
        let (index, read) = read_leb128_u32(&data[*pos..]);
        *pos += read;
        return vybe_runtime::opcode::heaptype::HeapType::Concrete(index);
    }
    let (value, read) = read_leb128_i32(&data[*pos..]);
    *pos += read;
    vybe_runtime::opcode::heaptype::HeapType::from_sleb(value)
}

fn skip_heaptype(data: &[u8], pos: &mut usize) {
    if data.get(*pos) == Some(&HEAPTYPE_EXACT) {
        *pos += 1;
        skip_leb128(data, pos); // u32 typeidx
        return;
    }
    skip_leb128(data, pos);
}

fn read_block_result_count(data: &[u8], pos: &mut usize) -> u8 {
    let first = data.get(*pos).copied().unwrap_or(0x40);
    skip_leb128(data, pos);
    if first == 0x40 { 0 } else { 1 }
}

fn read_leb128_i32(data: &[u8]) -> (i32, usize) {
    let mut result = 0i32;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        result |= ((byte & 0x7F) as i32) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 32 && (byte & 0x40 != 0) {
                result |= !0 << shift;
            }
            break;
        }
    }
    (result, pos)
}

fn decode_vybe_section(data: &[u8]) -> Result<Vec<Chunk>, String> {
    let mut pos = 0;
    let (name_len, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let _name = &data[pos..pos + name_len as usize];
    pos += name_len as usize;

    // Version byte
    let version = data[pos];
    pos += 1;

    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;

    let mut chunks = Vec::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize])
            .unwrap_or("")
            .to_string();
        pos += nlen as usize;
        let arity = data[pos];
        pos += 1;
        let (lc, read) = read_leb128_u32(&data[pos..]);
        pos += read;

        // Constants
        let (cc, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut constants = Vec::new();
        for _ in 0..cc {
            constants.push(decode_value(data, &mut pos));
        }

        // Imports
        let (ic, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut imports = Vec::new();
        for _ in 0..ic {
            let (mlen, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            let module = std::str::from_utf8(&data[pos..pos + mlen as usize])
                .unwrap_or("")
                .to_string();
            pos += mlen as usize;
            let (nlen, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            let iname = std::str::from_utf8(&data[pos..pos + nlen as usize])
                .unwrap_or("")
                .to_string();
            pos += nlen as usize;
            imports.push(vybe_runtime::chunk::Import {
                module,
                name: iname });
        }

        // Bytecode
        let (code_len, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let code = data[pos..pos + code_len as usize].to_vec();
        pos += code_len as usize;

        // Line info
        let (line_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut lines = Vec::with_capacity(line_count as usize);
        for _ in 0..line_count {
            let (line, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            lines.push(line);
        }

        // Type entries (v2+) with WASM Annotations format field descriptors
        let mut types = Vec::new();
        if version >= 2 {
            let (type_count, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            for _ in 0..type_count {
                // Type name
                let (nlen, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let type_name = std::str::from_utf8(&data[pos..pos + nlen as usize])
                    .unwrap_or("")
                    .to_string();
                pos += nlen as usize;

                // Declared supertype index (0 = no supertype)
                let (parent_index, read) = read_leb128_u32(&data[pos..]);
                pos += read;

                // Fields with descriptors
                let (field_count, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let mut fields = Vec::new();
                let mut field_descriptors = std::collections::HashMap::new();
                for _ in 0..field_count {
                    let (nlen, read) = read_leb128_u32(&data[pos..]);
                    pos += read;
                    let field_name = std::str::from_utf8(&data[pos..pos + nlen as usize])
                        .unwrap_or("")
                        .to_string();
                    pos += nlen as usize;
                    fields.push(field_name.clone());

                    // Decode field descriptor flags (WASM Annotations format)
                    let flags = data[pos];
                    pos += 1;
                    let descriptor = vybe_runtime::chunk::PropertyDescriptor {
                        writable: (flags & 0x01) != 0,
                        enumerable: (flags & 0x02) != 0,
                        configurable: (flags & 0x04) != 0 };
                    field_descriptors.insert(field_name, descriptor);
                }

                // Methods
                let (method_count, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let mut methods = Vec::new();
                for _ in 0..method_count {
                    let (nlen, read) = read_leb128_u32(&data[pos..]);
                    pos += read;
                    let method_name = std::str::from_utf8(&data[pos..pos + nlen as usize])
                        .unwrap_or("")
                        .to_string();
                    pos += nlen as usize;
                    let (chunk_idx, read) = read_leb128_u32(&data[pos..]);
                    pos += read;
                    methods.push((method_name, chunk_idx as usize));
                }

                // is_interface
                let is_interface = data[pos] != 0;
                pos += 1;

                // implements
                let (impl_count, read) = read_leb128_u32(&data[pos..]);
                pos += read;
                let mut implements = Vec::new();
                for _ in 0..impl_count {
                    let (iface_index, read) = read_leb128_u32(&data[pos..]);
                    pos += read;
                    implements.push(iface_index as u16);
                }

                // constructor_chunk
                let has_ctor = data[pos] != 0;
                pos += 1;
                let constructor_chunk = if has_ctor {
                    let (idx, read) = read_leb128_u32(&data[pos..]);
                    pos += read;
                    Some(idx as usize)
                } else {
                    None
                };

                types.push(vybe_runtime::chunk::TypeEntry {
                    name: type_name,
                    // The custom type section does not (yet) serialize the
                    // composite kind; decoded types are struct/class shaped.
                    // Array-kind round-tripping is part of the deferred
                    // imported-.wasm stamping work.
                    kind: vybe_runtime::chunk::CompositeKind::Struct,
                    parent_index: parent_index as u16,
                    fields,
                    methods,
                    is_interface,
                    implements,
                    constructor_chunk,
                    field_descriptors });
            }
        }

        let mut chunk = Chunk::new(&name);
        chunk.arity = arity;
        chunk.local_count = lc as u16;
        chunk.constants = constants;
        chunk.imports = imports;
        chunk.code = code;
        chunk.lines = lines;
        chunk.types = types;
        chunks.push(chunk);
    }
    Ok(chunks)
}

// ============================================================
// Helpers

// ============================================================
// Unit tests for the Custom Descriptors type-system encodings.
//
// Exact heap types and exact function imports exist only on the decode side —
// they have no representation in our bytecode, so they cannot be driven from
// an integration test through `write_wasm`. These exercise the section
// parsers directly.
// ============================================================
#[cfg(test)]
mod custom_descriptor_encoding_tests {
    use super::*;

    #[test]
    fn exact_heaptype_does_not_shift_later_type_indices() {
        // `heaptype ::= 0x62 x:u32 => exact x` is TWO lebs where every other
        // heaptype is one. A reader that consumes only the 0x62 goes on to
        // read the type index as the next value type, and the whole table
        // slides — which is how a single unread immediate corrupts every
        // function signature after it.
        //
        // type 0 = (func (param (ref null (exact 1))))
        // type 1 = (func (result i32))
        let mut section = vec![0x02];
        section.extend_from_slice(&[0x60, 0x01, 0x63, HEAPTYPE_EXACT, 0x01, 0x00]);
        section.extend_from_slice(&[0x60, 0x00, 0x01, 0x7F]);

        let parsed = parse_type_section(&section);
        assert_eq!(parsed.len(), 2, "both types must be recovered");
        assert_eq!(
            parsed[0],
            (vec![0x63u8], vec![]),
            "the exact heaptype's index must be consumed as part of the param"
        );
        assert_eq!(
            parsed[1],
            (vec![], vec![0x7Fu8]),
            "the type AFTER the exact heaptype must still be (func (result i32))"
        );
    }

    #[test]
    fn exact_heaptype_is_accepted_as_a_struct_field_type() {
        // (rec (type 0 (descriptor 1) (struct)) (type 1 (describes 0) (struct (field (ref (exact 0))))))
        // followed by a plain function type whose recovery proves the rec
        // group was consumed exactly.
        let mut section = vec![0x02];
        section.push(GC_REC);
        section.push(0x02);
        section.extend_from_slice(&[GC_SUB_FINAL, 0x00, CD_DESCRIPTOR, 0x01, GC_STRUCT, 0x00]);
        section.extend_from_slice(&[
            GC_SUB_FINAL,
            0x00,
            CD_DESCRIBES,
            0x00,
            GC_STRUCT,
            0x01,
            0x64,
            HEAPTYPE_EXACT,
            0x00,
            GC_IMMUT,
        ]);
        section.extend_from_slice(&[0x60, 0x00, 0x01, 0x7F]);

        let parsed = parse_type_section(&section);
        assert_eq!(
            parsed.len(),
            3,
            "the two struct types must each occupy an index"
        );
        assert_eq!(
            parsed[2],
            (vec![], vec![0x7Fu8]),
            "the function type after the rec group must land at index 2"
        );
    }

    #[test]
    fn gc_instructions_consume_exactly_their_spec_immediates() {
        // MVP.md §Instructions. Each entry is (sub-opcode, immediate bytes
        // that MUST be consumed). `struct.get`/`set` take typeidx + fieldidx;
        // `array.len` takes none; `array.new_fixed` takes typeidx + N. Getting
        // any of these wrong decodes the next instruction's first byte as an
        // opcode, which is a silent halt rather than a trap.
        let cases: &[(u32, &[u8])] = &[
            (0x00, &[0x03]),             // struct.new $t
            (0x01, &[0x03]),             // struct.new_default $t
            (0x02, &[0x03, 0x01]),       // struct.get $t i
            (0x03, &[0x03, 0x01]),       // struct.get_s $t i
            (0x04, &[0x03, 0x01]),       // struct.get_u $t i
            (0x05, &[0x03, 0x01]),       // struct.set $t i
            (0x06, &[0x03]),             // array.new $t
            (0x08, &[0x03, 0x02]),       // array.new_fixed $t N
            (0x0B, &[0x03]),             // array.get $t
            (0x0F, &[]),                 // array.len — NO immediate
            (0x10, &[0x03]),             // array.fill $t
            (0x11, &[0x03, 0x04]),       // array.copy $t1 $t2
            (0x20, &[0x03]),             // struct.new_desc $t
            (0x22, &[0x03]),             // ref.get_desc $t
        ];
        for (sub, immediates) in cases {
            // A trailing sentinel byte that must NOT be consumed.
            let mut bytes = immediates.to_vec();
            bytes.push(0xEE);
            let mut chunk = Chunk::new("<test>");
            let mut pos = 0usize;
            emit_gc_prefixed(&mut chunk, *sub, &bytes, &mut pos);
            assert_eq!(
                pos,
                immediates.len(),
                "0xFB {sub:#04x} must consume exactly {} immediate byte(s)",
                immediates.len()
            );
        }
    }

    #[test]
    fn exact_function_import_is_counted_as_a_function_import() {
        // `externtype ::= 0x20 x:typeidx => func exact x`. Miscounting it as
        // something other than a function import would shift the entire
        // function index space, so it must normalise to kind 0 and keep its
        // type index.
        let mut section = vec![0x01];
        section.extend_from_slice(&[0x01, b'm', 0x01, b'f']);
        section.push(EXTERNTYPE_FUNC_EXACT);
        section.push(0x07); // typeidx

        let details = parse_import_details(&section).expect("import section must parse");
        assert_eq!(details.len(), 1);
        assert_eq!(details[0].kind, 0, "an exact function import is a function import");
        assert_eq!(details[0].type_index, 7, "its type index must be read");

        let imports = parse_import_section(&section);
        assert_eq!(imports, vec![("m".to_string(), "f".to_string(), 0u8)]);
    }

    #[test]
    fn exact_function_type_in_the_export_section_is_malformed() {
        // "An export section using 0x20 is malfomed." — exports never declare
        // exactness; an exported function is exact iff its internal type is.
        let mut section = vec![0x01];
        section.extend_from_slice(&[0x01, b'f']);
        section.push(EXTERNTYPE_FUNC_EXACT);
        section.push(0x00);

        let err = validate_exports(&section, 1).expect_err("0x20 must be rejected");
        assert!(err.contains("exact function type"), "got: {err}");

        // The ordinary function export at the same index still validates.
        let mut ok = vec![0x01];
        ok.extend_from_slice(&[0x01, b'f']);
        ok.push(0x00);
        ok.push(0x00);
        validate_exports(&ok, 1).expect("a plain function export must still validate");
    }
}
