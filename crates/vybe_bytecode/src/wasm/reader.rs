//! WASM binary reader — decodes .wasm files into Chunk arrays.

use super::encoding::*;
use crate::value::Value;
use crate::{Chunk, Op};
use std::collections::HashSet;
use std::sync::Arc;

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
}

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
        &sections.memory_section,
        &sections.export_section,
        &sections.code_section,
    )
}

fn section_order_rank(section_id: u8) -> u8 {
    match section_id {
        SECTION_CUSTOM => 0,
        SECTION_TYPE => 1,
        SECTION_IMPORT => 2,
        SECTION_FUNCTION => 3,
        4 => 4, // table
        SECTION_MEMORY => 5,
        SECTION_GLOBAL => 6,
        SECTION_EXPORT => 7,
        8 => 8,  // start
        9 => 9,  // element
        12 => 10, // data_count is ordered before code
        SECTION_CODE => 11,
        11 => 12, // data
        SECTION_TAG => 13,
        other => other,
    }
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
    validate_start(&sections.start_section, &types, &imports, &func_type_indices)?;
    validate_element_section(&sections.elem_section, table_count)?;
    validate_data_sections(
        &sections.data_count_section,
        &sections.data_section,
        memory_count,
    )?;
    validate_code_bodies(
        &sections.code_section,
        &func_type_indices,
        &types,
        func_count,
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
    type_index: u32,
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
        let kind = data[pos];
        pos += 1;
        let (type_index, read) = read_leb128_u32(&data[pos..]);
        pos += read;
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
            let (table_idx, read) = read_leb128_u32(&data[pos..]);
            pos += read;
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
            let (memidx, read) = read_leb128_u32(&data_section[pos..]);
            pos += read;
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

fn validate_code_bodies(
    code_sec: &[u8],
    func_type_indices: &[u32],
    types: &[(Vec<u8>, Vec<u8>)],
    func_count: usize,
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
        let param_count = types
            .get(type_idx as usize)
            .map(|(params, _)| params.len())
            .ok_or_else(|| "Invalid WASM: function type index out of range".to_string())?;
        let total_locals = param_count + local_count;
        validate_instruction_stream(
            &code_sec[pos..body_end - 1],
            total_locals,
            func_count,
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

fn validate_instruction_stream(
    code: &[u8],
    local_count: usize,
    func_count: usize,
    global_count: usize,
    global_mutability: &[bool],
    data_count: usize,
    elem_count: usize,
    has_data_count_section: bool,
    uses_memory64: bool,
    uses_table64: bool,
) -> Result<(), String> {
    let mut pos = 0;
    let mut labels = Vec::<()>::new();
    let mut stack_depth = 0isize;
    while pos < code.len() {
        let op = code[pos];
        pos += 1;
        match op {
            0x01 => {}
            0x09 => {
                return Err(
                    "Invalid WASM: exception-handling rethrow is not decoded yet".into(),
                );
            }
            0x02 | 0x03 => {
                skip_leb128(code, &mut pos);
                labels.push(());
            }
            0x04 => {
                require_stack(&mut stack_depth, 1, "if condition")?;
                skip_leb128(code, &mut pos);
                labels.push(());
            }
            0x05 => {}
            0x0B => {
                labels.pop();
            }
            0x0C | 0x0D => {
                let (depth, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if depth as usize >= labels.len() {
                    return Err("Invalid WASM: branch depth out of range".into());
                }
                if op == 0x0D {
                    require_stack(&mut stack_depth, 1, "br_if condition")?;
                }
            }
            0x0E => {
                require_stack(&mut stack_depth, 1, "br_table selector")?;
                let (count, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                for _ in 0..count {
                    let (depth, read) = read_leb128_u32(&code[pos..]);
                    pos += read;
                    if depth as usize >= labels.len() {
                        return Err("Invalid WASM: br_table depth out of range".into());
                    }
                }
                let (default_depth, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if default_depth as usize >= labels.len() {
                    return Err("Invalid WASM: br_table default depth out of range".into());
                }
            }
            0x0F => {}
            0x10 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= func_count {
                    return Err("Invalid WASM: call function index out of range".into());
                }
            }
            0x1A => require_stack(&mut stack_depth, 1, "drop")?,
            0x1B => {
                require_stack(&mut stack_depth, 3, "select")?;
                stack_depth += 1;
            }
            0x18 => {
                return Err(
                    "Invalid WASM: exception-handling delegate is not decoded yet".into(),
                );
            }
            0x20 | 0x21 | 0x22 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= local_count {
                    return Err("Invalid WASM: local index out of range".into());
                }
                match op {
                    0x20 => stack_depth += 1,
                    0x21 => require_stack(&mut stack_depth, 1, "local.set")?,
                    0x22 => {}
                    _ => {}
                }
            }
            0x23 | 0x24 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= global_count {
                    return Err("Invalid WASM: global index out of range".into());
                }
                if op == 0x24 {
                    if !global_mutability.get(idx as usize).copied().unwrap_or(false) {
                        return Err("Invalid WASM: global.set to immutable global".into());
                    }
                    require_stack(&mut stack_depth, 1, "global.set")?;
                } else {
                    stack_depth += 1;
                }
            }
            0x25 | 0x26 if uses_table64 => {
                let name = if op == 0x25 { "table.get" } else { "table.set" };
                return Err(format!(
                    "Invalid WASM: table64 {name} is not decoded yet"
                ));
            }
            0x28..=0x40 => {
                if uses_memory64
                    && !matches!(op, 0x28 | 0x29 | 0x2B | 0x36 | 0x37 | 0x39 | 0x3F | 0x40)
                {
                    return Err(format!(
                        "Invalid WASM: memory64 {} is not decoded yet",
                        memory_opcode_name(op)
                    ));
                }
                let operands = if matches!(op, 0x36..=0x3E) { 2 } else if op == 0x40 { 1 } else { 1 };
                require_stack(&mut stack_depth, operands, "memory operation")?;
                if !matches!(op, 0x36..=0x3E) {
                    stack_depth += 1;
                }
                skip_memarg_or_memory_immediate(code, &mut pos, op);
            }
            0x41 => {
                skip_leb128(code, &mut pos);
                stack_depth += 1;
            }
            0x42 => {
                skip_leb128(code, &mut pos);
                stack_depth += 1;
            }
            0x43 => {
                pos += 4;
                stack_depth += 1;
            }
            0x44 => {
                pos += 8;
                stack_depth += 1;
            }
            0x45..=0x66 => {
                let operands = if op == 0x45 || op == 0x50 { 1 } else { 2 };
                require_stack(&mut stack_depth, operands, "comparison")?;
                stack_depth += 1;
            }
            0x67..=0xA6 => {
                let operands = if matches!(op, 0x67..=0x69 | 0x79..=0x7B | 0x8B..=0x91 | 0x99..=0x9F) {
                    1
                } else {
                    2
                };
                require_stack(&mut stack_depth, operands, "numeric operation")?;
                stack_depth += 1;
            }
            0xA7..=0xC4 => {}
            0xD0 => {
                pos += 1;
                stack_depth += 1;
            }
            0xD2 => {
                let (idx, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if idx as usize >= func_count {
                    return Err("Invalid WASM: ref.func index out of range".into());
                }
                stack_depth += 1;
            }
            0xE3 => {
                skip_leb128(code, &mut pos); // continuation type index
                let (handler_count, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                if handler_count != 0 {
                    return Err(
                        "Invalid WASM: resume handler vector is not decoded yet".into(),
                    );
                }
                require_stack(&mut stack_depth, 2, "resume")?;
            }
            0xFB => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
                    0x00..=0x01 | 0x06..=0x08 | 0x14..=0x19 => {
                        skip_leb128(code, &mut pos);
                    }
                    0x02..=0x05 | 0x09..=0x0A | 0x12..=0x13 => {
                        skip_leb128(code, &mut pos);
                        skip_leb128(code, &mut pos);
                    }
                    0x1C => {
                        require_stack(&mut stack_depth, 1, "ref.i31")?;
                        stack_depth += 1;
                    }
                    0x1D..=0x1E => {}
                    _ => {}
                }
            }
            0xFC => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
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
                        require_stack(&mut stack_depth, 3, "memory.init")?;
                    }
                    0x09 => {
                        let (data_idx, read) = read_leb128_u32(&code[pos..]);
                        pos += read;
                        if data_idx as usize >= data_count {
                            return Err("Invalid WASM: data.drop index out of range".into());
                        }
                    }
                    0x0C => {
                        if uses_table64 {
                            return Err(
                                "Invalid WASM: table64 table.init is not decoded yet".into(),
                            );
                        }
                        let (elem_idx, read) = read_leb128_u32(&code[pos..]);
                        pos += read;
                        if elem_idx as usize >= elem_count {
                            return Err("Invalid WASM: table.init element index out of range".into());
                        }
                        skip_leb128(code, &mut pos);
                        require_stack(&mut stack_depth, 3, "table.init")?;
                    }
                    0x0A | 0x0B if uses_memory64 => {
                        let name = if sub == 0x0A { "memory.copy" } else { "memory.fill" };
                        return Err(format!(
                            "Invalid WASM: memory64 {name} is not decoded yet"
                        ));
                    }
                    0x0E | 0x0F | 0x10 | 0x11 if uses_table64 => {
                        return Err(format!(
                            "Invalid WASM: table64 {} is not decoded yet",
                            table_opcode_name(sub)
                        ));
                    }
                    _ => {}
                }
            }
            0xFD => {
                let (sub, read) = read_leb128_u32(&code[pos..]);
                pos += read;
                match sub {
                    0x00..=0x0B | 0x54..=0x5D => {
                        if uses_memory64 {
                            return Err(format!(
                                "Invalid WASM: memory64 {} is not decoded yet",
                                simd_memory_opcode_name(sub)
                            ));
                        }
                        skip_memarg(code, &mut pos);
                        if sub == 0x0B || (0x58..=0x5B).contains(&sub) {
                            require_stack(&mut stack_depth, 2, "simd memory store")?;
                        } else {
                            require_stack(&mut stack_depth, 1, "simd memory load")?;
                            stack_depth += 1;
                        }
                        if (0x54..=0x5B).contains(&sub) {
                            pos = pos.saturating_add(1).min(code.len());
                        }
                    }
                    0x0C => {
                        pos = pos.saturating_add(16).min(code.len());
                        stack_depth += 1;
                    }
                    0x0D => {
                        pos = pos.saturating_add(16).min(code.len());
                        require_stack(&mut stack_depth, 2, "i8x16.shuffle")?;
                        stack_depth += 1;
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
                if uses_memory64 && matches!(sub, 0x00..=0x02 | 0x10..=0x4E) {
                    return Err(format!(
                        "Invalid WASM: memory64 {} is not decoded yet",
                        atomic_opcode_name(sub)
                    ));
                }
                match sub {
                    0x00 => {
                        skip_memarg(code, &mut pos);
                        require_stack(&mut stack_depth, 2, "memory.atomic.notify")?;
                        stack_depth += 1;
                    }
                    0x01 | 0x02 => {
                        skip_memarg(code, &mut pos);
                        require_stack(&mut stack_depth, 3, "memory.atomic.wait")?;
                        stack_depth += 1;
                    }
                    0x03 => {
                        pos = pos.saturating_add(1).min(code.len());
                    }
                    0x10..=0x16 => {
                        skip_memarg(code, &mut pos);
                        require_stack(&mut stack_depth, 1, "atomic load")?;
                        stack_depth += 1;
                    }
                    0x17..=0x1D => {
                        skip_memarg(code, &mut pos);
                        require_stack(&mut stack_depth, 2, "atomic store")?;
                    }
                    0x1E..=0x47 => {
                        skip_memarg(code, &mut pos);
                        require_stack(&mut stack_depth, 2, "atomic rmw")?;
                        stack_depth += 1;
                    }
                    0x48..=0x4E => {
                        skip_memarg(code, &mut pos);
                        require_stack(&mut stack_depth, 3, "atomic cmpxchg")?;
                        stack_depth += 1;
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn require_stack(depth: &mut isize, n: isize, context: &str) -> Result<(), String> {
    if *depth < n {
        return Err(format!("Invalid WASM: stack underflow in {context}"));
    }
    *depth -= n;
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
    memory_sec: &[u8],
    export_sec: &[u8],
    code_sec: &[u8],
) -> Result<Vec<Chunk>, String> {
    // Parse type section to get function signatures
    let types = parse_type_section(type_sec);
    let func_type_indices = parse_function_section(func_sec);

    // Parse imports
    let imports = parse_import_section(import_sec);
    let import_func_count = imports.iter().filter(|(_, _, kind)| *kind == 0).count();

    // Parse exports to find function names
    let exports = parse_export_section(export_sec);
    let memory_min_pages = parse_memory_section(memory_sec);
    let uses_memory64 = section_uses_memory64(memory_sec);

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
        );
        chunk.result_arity = result_arity;
        chunk.memory_min_pages = memory_min_pages.clone();
        chunk.emit_op(Op::RETURN, 0);
        chunks.push(chunk);

        cpos = body_end;
    }

    // Add imports to script chunk
    for (module, name, _) in &imports {
        script.add_import(module, name);
    }
    script.memory_min_pages = memory_min_pages;

    // Insert script as chunk 0
    chunks.insert(0, script);

    Ok(chunks)
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
) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk.local_count = arity as u16 + wasm_local_count as u16;

    let mut pos = 0;
    let mut label_stack: Vec<()> = Vec::new();

    while pos < wasm.len() {
        let byte = wasm[pos];
        pos += 1;

        match byte {
            0x00 => chunk.emit_op(Op::HALT, 0),
            0x01 => {} // nop

            // block blocktype — forward jump target
            0x02 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                chunk.emit_block_typed(0, result_count);
                label_stack.push(());
            }

            // loop blocktype — backward jump target
            0x03 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                chunk.emit_loop_typed(0, result_count);
                label_stack.push(());
            }

            // if blocktype — conditional block
            0x04 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                if result_count == 0 {
                    chunk.emit_if(0);
                } else {
                    chunk.emit_if_value(0);
                }
                label_stack.push(());
            }

            // else
            0x05 => {
                chunk.emit_else(0);
            }

            // end
            0x0B => {
                let _ = label_stack.pop();
                chunk.emit_end(0);
            }

            // br N — branch to Nth enclosing label
            0x0C => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br(depth, 0);
            }

            // br_if N — conditional branch
            0x0D => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br_if(depth, 0);
            }
            0x0E => {
                // br_table — branch table
                let (count, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let mut depths = Vec::with_capacity(count as usize);
                for _ in 0..count {
                    let (depth, _) = read_leb128_u32(&wasm[pos..]);
                    skip_leb128(wasm, &mut pos);
                    depths.push(depth);
                }
                let (default_depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br_table(&depths, default_depth, 0);
            }
            0x0F => chunk.emit_op(Op::RETURN, 0),
            0x1A => chunk.emit_op(Op::DROP, 0),
            0x1B => chunk.emit_op(Op::SELECT, 0),

            // call — adjust index (skip imports, offset to our chunk indices)
            0x10 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u8(Op::CALL, idx as u8, 0);
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
            // local.tee
            0x22 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_SET, idx as u16, 0);
            }

            // i32.const
            0x41 => {
                let (val, read) = read_leb128_i32(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::I32(val));
                chunk.emit_op_u16(Op::CONST, ci, 0);
            }

            // i64.const
            0x42 => {
                let (val, read) = read_leb128_i64(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::I64(val));
                chunk.emit_op_u16(Op::CONST, ci, 0);
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
                    let ci = chunk.add_constant(Value::F64(val));
                    chunk.emit_op_u16(Op::CONST, ci, 0);
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
                    let ci = chunk.add_constant(Value::F64(val as f64));
                    chunk.emit_op_u16(Op::CONST, ci, 0);
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
                chunk.emit_op(if uses_memory64 { Op::I32_LOAD_64 } else { Op::I32_LOAD }, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x29 => {
                chunk.emit_op(if uses_memory64 { Op::I64_LOAD_64 } else { Op::I64_LOAD }, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2A => {
                chunk.emit_op(Op::F32_LOAD, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x2B => {
                chunk.emit_op(if uses_memory64 { Op::F64_LOAD_64 } else { Op::F64_LOAD }, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x2C => {
                chunk.emit_op(Op::I32_LOAD8_S, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x2D => {
                chunk.emit_op(Op::I32_LOAD8_U, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x2E => {
                chunk.emit_op(Op::I32_LOAD16_S, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x2F => {
                chunk.emit_op(Op::I32_LOAD16_U, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x30 => {
                chunk.emit_op(Op::I64_LOAD8_S, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x31 => {
                chunk.emit_op(Op::I64_LOAD8_U, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x32 => {
                chunk.emit_op(Op::I64_LOAD16_S, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x33 => {
                chunk.emit_op(Op::I64_LOAD16_U, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x34 => {
                chunk.emit_op(Op::I64_LOAD32_S, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x35 => {
                chunk.emit_op(Op::I64_LOAD32_U, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x36 => {
                chunk.emit_op(if uses_memory64 { Op::I32_STORE_64 } else { Op::I32_STORE }, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x37 => {
                chunk.emit_op(if uses_memory64 { Op::I64_STORE_64 } else { Op::I64_STORE }, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x38 => {
                chunk.emit_op(Op::F32_STORE, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x39 => {
                chunk.emit_op(if uses_memory64 { Op::F64_STORE_64 } else { Op::F64_STORE }, 0);
                read_emit_memarg_for_memory_width(&mut chunk, wasm, &mut pos, uses_memory64);
            }
            0x3A => {
                chunk.emit_op(Op::I32_STORE8, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x3B => {
                chunk.emit_op(Op::I32_STORE16, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x3C => {
                chunk.emit_op(Op::I64_STORE8, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x3D => {
                chunk.emit_op(Op::I64_STORE16, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x3E => {
                chunk.emit_op(Op::I64_STORE32, 0);
                read_emit_memarg(&mut chunk, wasm, &mut pos);
            }
            0x3F => {
                chunk.emit_op(
                    if uses_memory64 {
                        Op::I64_MEMORY_SIZE
                    } else {
                        Op::MEMORY_SIZE
                    },
                    0,
                );
                read_emit_leb_u32(&mut chunk, wasm, &mut pos);
            }
            0x40 => {
                chunk.emit_op(
                    if uses_memory64 {
                        Op::I64_MEMORY_GROW
                    } else {
                        Op::MEMORY_GROW
                    },
                    0,
                );
                read_emit_leb_u32(&mut chunk, wasm, &mut pos);
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

            // global.get/set — WASM globals mapped to global_get/set with index as name
            0x23 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let name = format!("__wasm_global_{}", idx);
                let ci = chunk.add_constant(Value::String(Arc::from(name.as_str())));
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
            0x24 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let name = format!("__wasm_global_{}", idx);
                let ci = chunk.add_constant(Value::String(Arc::from(name.as_str())));
                chunk.emit_op_u16(Op::GLOBAL_SET, ci, 0);
            }

            // call_indirect
            0x11 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
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
                        read_emit_leb_u32(&mut chunk, wasm, &mut pos);
                    }
                    0x09 => {
                        let (data_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::DATA_DROP, data_idx as u8, 0);
                    }
                    0x0A => {
                        chunk.emit_op(Op::MEMORY_COPY, 0);
                        read_emit_leb_u32(&mut chunk, wasm, &mut pos); // dst memory
                        read_emit_leb_u32(&mut chunk, wasm, &mut pos); // src memory
                    }
                    0x0B => {
                        chunk.emit_op(Op::MEMORY_FILL, 0);
                        read_emit_leb_u32(&mut chunk, wasm, &mut pos);
                    }
                    0x0C => {
                        let (elem_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        skip_leb128(wasm, &mut pos); // table index
                        chunk.emit_op_u8(Op::TABLE_INIT, elem_idx as u8, 0);
                    }
                    0x0D => {
                        let (elem_idx, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        chunk.emit_op_u8(Op::ELEM_DROP, elem_idx as u8, 0);
                    }
                    0x0E => {
                        let (dst_table, _) = read_leb128_u32(&wasm[pos..]);
                        skip_leb128(wasm, &mut pos);
                        skip_leb128(wasm, &mut pos); // src table
                        chunk.emit_op_u8(Op::TABLE_COPY, dst_table as u8, 0);
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

            // SIMD and relaxed-SIMD proposal prefix.
            0xFD => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_simd_prefixed(&mut chunk, sub, wasm, &mut pos);
            }

            // Threads/atomics proposal prefix.
            0xFE => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_threads_prefixed(&mut chunk, sub, wasm, &mut pos);
            }

            // Unknown — skip
            _ => {}
        }
    }

    chunk
}

fn emit_gc_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    let Some(op) = u8::try_from(sub).ok().and_then(|s| Op::decode(0xFB, s)) else {
        return;
    };
    match op {
        _ if op == Op::STRUCT_NEW
            || op == Op::STRUCT_NEW_DEFAULT
            || op == Op::STRUCT_GET
            || op == Op::STRUCT_GET_S
            || op == Op::STRUCT_GET_U
            || op == Op::STRUCT_SET
            || op == Op::ARRAY_NEW
            || op == Op::ARRAY_NEW_DEFAULT
            || op == Op::ARRAY_GET
            || op == Op::ARRAY_GET_S
            || op == Op::ARRAY_GET_U
            || op == Op::ARRAY_SET
            || op == Op::ARRAY_LENGTH
            || op == Op::ARRAY_FILL =>
        {
            let (idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            match op.operand_format() {
                crate::opcode::OperandFormat::U16 => chunk.emit_op_u16(op, idx as u16, 0),
                _ => chunk.emit_op(op, 0),
            }
        }
        _ if op == Op::ARRAY_NEW_FIXED => {
            let (_type_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            let (extra, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_op_u16(op, extra as u16, 0);
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
            skip_heaptype(wasm, pos);
            let idx = chunk.add_constant(Value::String(Arc::from("__wasm_heaptype")));
            chunk.emit_op_u16(op, idx, 0);
        }
        _ if op == Op::BR_ON_CAST || op == Op::BR_ON_CAST_FAIL => {
            skip_leb128(wasm, pos); // flags
            let (depth, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            skip_heaptype(wasm, pos);
            skip_heaptype(wasm, pos);
            let idx = chunk.add_constant(Value::String(Arc::from("__wasm_heaptype")));
            chunk.emit_op_u16(op, idx, 0);
            chunk.emit(depth as u8, 0);
        }
        _ => chunk.emit_op(op, 0),
    }
}

fn emit_simd_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    if (0x100..=0x113).contains(&sub) {
        let relaxed_sub = (sub - 0x100) as u8;
        if let Some(op) = Op::decode(0xDD, relaxed_sub) {
            chunk.emit_op(op, 0);
        }
        return;
    }

    let Some(op) = u8::try_from(sub).ok().and_then(|s| Op::decode(0xFD, s)) else {
        return;
    };
    match op {
        _ if op == Op::V128_LOAD || op == Op::V128_STORE => {
            skip_memarg(wasm, pos);
            chunk.emit_op(op, 0);
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
        _ if op.operand_format() == crate::opcode::OperandFormat::U8 => {
            let lane = wasm.get(*pos).copied().unwrap_or(0);
            *pos += 1;
            chunk.emit_op_u8(op, lane, 0);
        }
        _ => chunk.emit_op(op, 0),
    }
}

fn emit_threads_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    let Some(op) = u8::try_from(sub).ok().and_then(|s| Op::decode(0xFE, s)) else {
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
            copy_memarg(wasm, pos, chunk);
        }
    }
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

fn parse_type_section(data: &[u8]) -> Vec<(Vec<u8>, Vec<u8>)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut types = Vec::new();
    for _ in 0..count {
        if pos >= data.len() || data[pos] != TYPE_FUNC {
            pos += 1;
            continue;
        }
        pos += 1; // skip 0x60
        let (param_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let params: Vec<u8> = data[pos..pos + param_count as usize].to_vec();
        pos += param_count as usize;
        let (result_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let results: Vec<u8> = data[pos..pos + result_count as usize].to_vec();
        pos += result_count as usize;
        types.push((params, results));
    }
    types
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
        let kind = data[pos];
        pos += 1;
        skip_leb128(&data, &mut pos); // type index or other descriptor
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

fn read_emit_leb_u32(chunk: &mut Chunk, data: &[u8], pos: &mut usize) -> u32 {
    let (value, read) = read_leb128_u32(&data[*pos..]);
    for byte in &data[*pos..*pos + read] {
        chunk.emit(*byte, 0);
    }
    *pos += read;
    value
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

fn memory_opcode_name(op: u8) -> &'static str {
    match op {
        0x2A => "f32.load",
        0x2C => "i32.load8_s",
        0x2D => "i32.load8_u",
        0x2E => "i32.load16_s",
        0x2F => "i32.load16_u",
        0x30 => "i64.load8_s",
        0x31 => "i64.load8_u",
        0x32 => "i64.load16_s",
        0x33 => "i64.load16_u",
        0x34 => "i64.load32_s",
        0x35 => "i64.load32_u",
        0x38 => "f32.store",
        0x3A => "i32.store8",
        0x3B => "i32.store16",
        0x3C => "i64.store8",
        0x3D => "i64.store16",
        0x3E => "i64.store32",
        _ => "memory operation",
    }
}

fn table_opcode_name(sub: u32) -> &'static str {
    match sub {
        0x0C => "table.init",
        0x0E => "table.copy",
        0x0F => "table.grow",
        0x10 => "table.size",
        0x11 => "table.fill",
        _ => "table operation",
    }
}

fn simd_memory_opcode_name(sub: u32) -> &'static str {
    match sub {
        0x00 => "v128.load",
        0x0B => "v128.store",
        _ => "SIMD memory operation",
    }
}

fn atomic_opcode_name(sub: u32) -> &'static str {
    match sub {
        0x10 => "i32.atomic.load",
        0x17 => "i32.atomic.store",
        _ => "atomic memory operation",
    }
}

fn skip_memarg(data: &[u8], pos: &mut usize) {
    skip_leb128(data, pos); // align
    skip_leb128(data, pos); // offset
}

fn skip_heaptype(data: &[u8], pos: &mut usize) {
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
    let _version = data[pos];
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
            imports.push(crate::chunk::Import {
                module,
                name: iname,
            });
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

        let mut chunk = Chunk::new(&name);
        chunk.arity = arity;
        chunk.local_count = lc as u16;
        chunk.constants = constants;
        chunk.imports = imports;
        chunk.code = code;
        chunk.lines = lines;
        chunks.push(chunk);
    }
    Ok(chunks)
}

// ============================================================
// Helpers
