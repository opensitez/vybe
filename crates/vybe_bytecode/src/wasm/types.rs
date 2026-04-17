//! WASM type section encoding.
//! Emits GC struct types from TypeEntry, array type, and function types.

use super::encoding::*;
use crate::Chunk;

/// Context for .wasm emission — maps internal types to WASM type indices.
pub struct WasmTypeContext {
    /// type_name (lowercased) → WASM type index for GC struct types
    pub struct_type_indices: std::collections::HashMap<String, u32>,
    /// type_name → vec of field names in order (for field index lookup)
    pub struct_fields: std::collections::HashMap<String, Vec<String>>,
    /// WASM type index for the dynamic array type
    pub array_type_idx: u32,
    /// First function type index (after GC types)
    pub func_type_base: u32,
}

impl WasmTypeContext {
    /// Look up the WASM type index for a struct type by name.
    pub fn struct_type(&self, name: &str) -> Option<u32> {
        self.struct_type_indices.get(&name.to_lowercase()).copied()
    }

    /// Look up the field index for a field name within a struct type.
    pub fn field_index(&self, type_name: &str, field_name: &str) -> Option<u32> {
        let fields = self.struct_fields.get(&type_name.to_lowercase())?;
        fields.iter().position(|f| f == &field_name.to_lowercase()).map(|i| i as u32)
    }
}

/// Build the type context and encode the type section.
/// Returns (encoded_bytes, context).
pub fn build_type_context(chunks: &[Chunk], import_count: usize) -> (Vec<u8>, WasmTypeContext) {
    let mut out = Vec::new();
    let mut ctx = WasmTypeContext {
        struct_type_indices: std::collections::HashMap::new(),
        struct_fields: std::collections::HashMap::new(),
        array_type_idx: 0,
        func_type_base: 0,
    };

    // Collect TypeEntry definitions from chunk 0
    let type_entries: Vec<&crate::chunk::TypeEntry> = chunks.first()
        .map(|c| c.types.iter().collect())
        .unwrap_or_default();

    // Layout: [GC struct types...] [array type] [function types...]
    let gc_struct_count = type_entries.len() as u32;
    let array_count = 1u32;
    let func_count = (import_count + chunks.len()) as u32;
    let total = gc_struct_count + array_count + func_count;

    write_leb128_u32(&mut out, total);

    // ── GC struct types ──
    for (i, te) in type_entries.iter().enumerate() {
        let type_idx = i as u32;
        let name_lower = te.name.to_lowercase();
        ctx.struct_type_indices.insert(name_lower.clone(), type_idx);
        ctx.struct_fields.insert(name_lower, te.fields.clone());

        // Emit struct type: 0x5F field_count (field)*
        out.push(GC_STRUCT);
        write_leb128_u32(&mut out, te.fields.len() as u32);
        for _ in &te.fields {
            out.push(TYPE_EXTERNREF); // each field holds an externref
            out.push(GC_MUT);
        }
    }

    // ── Array type ──
    ctx.array_type_idx = gc_struct_count;
    out.push(GC_ARRAY);
    out.push(TYPE_EXTERNREF);
    out.push(GC_MUT);

    // ── Function types ──
    ctx.func_type_base = gc_struct_count + array_count;

    // Import function types
    for _ in 0..import_count {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 0);
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_EXTERNREF);
    }

    // Chunk function types
    for chunk in chunks {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, chunk.arity as u32);
        for _ in 0..chunk.arity { out.push(TYPE_EXTERNREF); }
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_EXTERNREF);
    }

    (out, ctx)
}
