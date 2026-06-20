use vybe_bytecode::Chunk;

/// Lua-specific emit dispatcher for `common:lua.*` operations.
/// Returns `true` if the name was handled, `false` if not (so caller can fall through).
pub fn dispatch(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    // Strip the "lua." prefix to get the actual operation
    let op = if let Some(stripped) = name.strip_prefix("lua.") {
        stripped
    } else {
        return false;
    };

    match op {
        // Metatable operations
        "metatable_index" => emit_metatable_index(chunks, current, argc, line),
        "metatable_newindex" => emit_metatable_newindex(chunks, current, argc, line),
        "metatable_call" => emit_metatable_call(chunks, current, argc, line),
        "metatable_add" => emit_metatable_add(chunks, current, argc, line),
        "metatable_sub" => emit_metatable_sub(chunks, current, argc, line),
        "metatable_mul" => emit_metatable_mul(chunks, current, argc, line),
        "metatable_div" => emit_metatable_div(chunks, current, argc, line),
        "metatable_mod" => emit_metatable_mod(chunks, current, argc, line),
        "metatable_pow" => emit_metatable_pow(chunks, current, argc, line),
        "metatable_unm" => emit_metatable_unm(chunks, current, argc, line),
        "metatable_concat" => emit_metatable_concat(chunks, current, argc, line),
        "metatable_eq" => emit_metatable_eq(chunks, current, argc, line),
        "metatable_lt" => emit_metatable_lt(chunks, current, argc, line),
        "metatable_le" => emit_metatable_le(chunks, current, argc, line),
        "metatable_len" => emit_metatable_len(chunks, current, argc, line),
        
        // Standard library modules
        "io" => emit_io_module(chunks, current, argc, line),
        "os" => emit_os_module(chunks, current, argc, line),
        "string" => emit_string_module(chunks, current, argc, line),
        "table" => emit_table_module(chunks, current, argc, line),
        "math" => emit_math_module(chunks, current, argc, line),
        "debug" => emit_debug_module(chunks, current, argc, line),
        "package" => emit_package_module(chunks, current, argc, line),
        "coroutine" => emit_coroutine_module(chunks, current, argc, line),
        
        _ => return false,
    }
    
    true
}

fn emit_metatable_index(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __index metamethod: table[key] when key not found
    // For now, emit a generic property access that will work with JS objects
    // TODO: Implement proper Lua metatable semantics
}

fn emit_metatable_newindex(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __newindex metamethod: table[key] = value when key not found
    // TODO: Implement proper Lua metatable semantics
}

fn emit_metatable_call(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __call metamethod: table(...) when table is called as a function
    // TODO: Implement proper Lua metatable semantics
}

fn emit_metatable_add(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __add metamethod: table + other
    // Emit generic addition, will work with JS numbers
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_sub(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __sub metamethod: table - other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_mul(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __mul metamethod: table * other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_div(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __div metamethod: table / other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_mod(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __mod metamethod: table % other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_pow(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __pow metamethod: table ^ other
    // Lua uses ^ for exponentiation
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_unm(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __unm metamethod: -table
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_concat(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __concat metamethod: table .. other
    // Lua uses .. for concatenation
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_eq(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __eq metamethod: table == other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_lt(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __lt metamethod: table < other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_le(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __le metamethod: table <= other
    // TODO: Implement Lua metatable dispatch
}

fn emit_metatable_len(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // __len metamethod: #table
    // Lua uses # for length operator
    // TODO: Implement Lua metatable dispatch
}

fn emit_io_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua io module - map to JS console/window APIs
    // For now, just emit a placeholder
}

fn emit_os_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua os module - map to JS Date/performance APIs
}

fn emit_string_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua string module - already mapped to JS String methods in profile
}

fn emit_table_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua table module - map to JS Array/Object methods
}

fn emit_math_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua math module - already mapped to JS Math in profile
}

fn emit_debug_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua debug module - map to JS debugger/console APIs
}

fn emit_package_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua package module - Lua's module system
    // For "lua over js", we can map to JS require/import
}

fn emit_coroutine_module(
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) {
    // Lua coroutine module - map to JS generators/async functions
    // Lua coroutines can be implemented using JS generators
}