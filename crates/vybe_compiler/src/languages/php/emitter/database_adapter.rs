use std::sync::Arc;

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

use crate::emitter::collections;

const PDO_FETCH_COLUMN: f64 = 7.0;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, value: Value, line: u32) {
    let idx = chunk.add_constant(value);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[0].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn string_slot_nonempty(chunk: &mut Chunk, slot: u16, line: u32) -> usize {
    lget(chunk, slot, line);
    push_str(chunk, "", line);
    chunk.emit_op(Op::DYN_NE, line);
    chunk.emit_jump(Op::BR_IF_FALSE, line)
}

fn concat_slot_with_literal(chunk: &mut Chunk, slot: u16, suffix: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, suffix, line);
    chunk.emit_op(Op::STR_CONCAT, line);
    lset(chunk, slot, line);
}

fn concat_slot_with_slot(chunk: &mut Chunk, target_slot: u16, value_slot: u16, line: u32) {
    lget(chunk, target_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::STR_CONCAT, line);
    lset(chunk, target_slot, line);
}

fn replace_in_slot(chunk: &mut Chunk, slot: u16, from: &str, to: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, from, line);
    push_str(chunk, to, line);
    chunk.emit_op(Op::STR_REPLACE, line);
    lset(chunk, slot, line);
}

fn append_credentials(chunk: &mut Chunk, normalized_slot: u16, username_slot: Option<u16>, password_slot: Option<u16>, line: u32) {
    if let Some(slot) = username_slot {
        let skip = string_slot_nonempty(chunk, slot, line);
        concat_slot_with_literal(chunk, normalized_slot, ";user=", line);
        concat_slot_with_slot(chunk, normalized_slot, slot, line);
        chunk.patch_jump(skip);
    }

    if let Some(slot) = password_slot {
        let skip = string_slot_nonempty(chunk, slot, line);
        concat_slot_with_literal(chunk, normalized_slot, ";password=", line);
        concat_slot_with_slot(chunk, normalized_slot, slot, line);
        chunk.patch_jump(skip);
    }
}

fn normalize_pdo_dsn(
    chunks: &mut [Chunk],
    current: usize,
    dsn_slot: u16,
    username_slot: Option<u16>,
    password_slot: Option<u16>,
    line: u32,
) -> u16 {
    let chunk = &mut chunks[current];
    let normalized_slot = alloc_local(chunk);
    lget(chunk, dsn_slot, line);
    lset(chunk, normalized_slot, line);

    lget(chunk, dsn_slot, line);
    push_str(chunk, "mysql:", line);
    chunk.emit_op(Op::STR_STARTS_WITH, line);
    let not_mysql = chunk.emit_jump(Op::BR_IF_FALSE, line);

    replace_in_slot(chunk, normalized_slot, "mysql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);
    let done = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(not_mysql);
    lget(chunk, dsn_slot, line);
    push_str(chunk, "pgsql:", line);
    chunk.emit_op(Op::STR_STARTS_WITH, line);
    let not_pgsql = chunk.emit_jump(Op::BR_IF_FALSE, line);

    replace_in_slot(chunk, normalized_slot, "pgsql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    chunk.emit_op(Op::STR_CONTAINS, line);
    let has_port = chunk.emit_jump(Op::BR_IF_TRUE, line);
    concat_slot_with_literal(chunk, normalized_slot, ";port=5432", line);
    chunk.patch_jump(has_port);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);
    let done_pgsql = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(not_pgsql);
    lget(chunk, dsn_slot, line);
    push_str(chunk, "postgres:", line);
    chunk.emit_op(Op::STR_STARTS_WITH, line);
    let not_postgres = chunk.emit_jump(Op::BR_IF_FALSE, line);

    replace_in_slot(chunk, normalized_slot, "postgres:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    chunk.emit_op(Op::STR_CONTAINS, line);
    let has_postgres_port = chunk.emit_jump(Op::BR_IF_TRUE, line);
    concat_slot_with_literal(chunk, normalized_slot, ";port=5432", line);
    chunk.patch_jump(has_postgres_port);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);

    chunk.patch_jump(done);
    chunk.patch_jump(done_pgsql);
    chunk.patch_jump(not_postgres);
    normalized_slot
}

fn stamp_pdo_type(chunk: &mut Chunk, conn_slot: u16, line: u32) {
    lget(chunk, conn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let skip = chunk.emit_jump(Op::BR_IF_TRUE, line);
    lget(chunk, conn_slot, line);
    push_str(chunk, "PDO", line);
    struct_set_key(chunk, "__type", line);
    chunk.patch_jump(skip);
}

fn emit_first_column_value(chunks: &mut [Chunk], current: usize, row_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    lget(chunk, row_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let keys_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, row_slot, line);
    lget(chunk, keys_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_empty_array(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

fn emit_new_statement(
    chunks: &mut [Chunk],
    current: usize,
    conn_slot: u16,
    sql_slot: Option<u16>,
    rows_slot: Option<u16>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, conn_slot, line);
    call_import(chunks, current, "vybe:database", "createCommand", 1, line);
    let stmt_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    push_str(chunk, "PDOStatement", line);
    struct_set_key(chunk, "__type", line);

    if let Some(slot) = sql_slot {
        lget(chunk, stmt_slot, line);
        lget(chunk, slot, line);
        struct_set_key(chunk, "commandtext", line);
    }

    lget(chunk, stmt_slot, line);
    if let Some(slot) = rows_slot {
        lget(chunk, slot, line);
    } else {
        emit_empty_array(chunks, current, line);
    }
    let chunk = &mut chunks[current];
    struct_set_key(chunk, "__rows", line);

    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);

    lget(chunk, stmt_slot, line);
}

fn emit_queryish_check(chunk: &mut Chunk, sql_slot: u16, prefix: &str, line: u32) -> usize {
    lget(chunk, sql_slot, line);
    push_str(chunk, prefix, line);
    chunk.emit_op(Op::STR_STARTS_WITH, line);
    chunk.emit_jump(Op::BR_IF_TRUE, line)
}

pub fn emit_php_pdo_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let options_slot = if argc >= 4 { Some(alloc_local(chunk)) } else { None };
    let password_slot = if argc >= 3 { Some(alloc_local(chunk)) } else { None };
    let username_slot = if argc >= 2 { Some(alloc_local(chunk)) } else { None };
    let dsn_slot = alloc_local(chunk);

    if let Some(slot) = options_slot { lset(chunk, slot, line); }
    if let Some(slot) = password_slot { lset(chunk, slot, line); }
    if let Some(slot) = username_slot { lset(chunk, slot, line); }
    lset(chunk, dsn_slot, line);

    let normalized_slot = normalize_pdo_dsn(chunks, current, dsn_slot, username_slot, password_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, normalized_slot, line);
    call_import(chunks, current, "vybe:database", "connect", 1, line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);
    stamp_pdo_type(chunk, conn_slot, line);
    lget(chunk, conn_slot, line);
}

pub fn emit_php_pdo_query(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, sql_slot, line);
    call_import(chunks, current, "vybe:database", "query", 2, line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    emit_new_statement(chunks, current, conn_slot, Some(sql_slot), Some(rows_slot), line);
}

pub fn emit_php_pdo_exec(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, sql_slot, line);
    call_import(chunks, current, "vybe:database", "execute", 2, line);
}

pub fn emit_php_pdo_prepare(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    emit_new_statement(chunks, current, conn_slot, Some(sql_slot), None, line);
}

pub fn emit_php_pdo_set_attribute(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let attr_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, value_slot, line);
    lset(chunk, attr_slot, line);
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, "__pdo_attr", line);
    chunk.emit_op(Op::TRUE, line);
}

pub fn emit_php_pdo_begin_transaction(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "vybe:database", "beginTransaction", 1, line);
}

pub fn emit_php_pdo_commit(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "vybe:database", "commit", 1, line);
}

pub fn emit_php_pdo_rollback(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "vybe:database", "rollback", 1, line);
}

pub fn emit_php_pdo_statement_execute(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let params_slot = if argc >= 1 { Some(alloc_local(chunk)) } else { None };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = params_slot { lset(chunk, slot, line); }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "commandtext", line);
    chunk.emit_op(Op::STR_TRIM, line);
    chunk.emit_op(Op::STR_TO_LOWER, line);
    let sql_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, sql_slot, line);

    let query_jumps = vec![
        emit_queryish_check(chunk, sql_slot, "select", line),
        emit_queryish_check(chunk, sql_slot, "pragma", line),
        emit_queryish_check(chunk, sql_slot, "show", line),
        emit_queryish_check(chunk, sql_slot, "with", line),
        emit_queryish_check(chunk, sql_slot, "describe", line),
    ];
    let not_query = chunk.emit_jump(Op::BR, line);

    for jump in query_jumps {
        chunk.patch_jump(jump);
    }
    lget(chunk, stmt_slot, line);
    if let Some(slot) = params_slot {
        chunk.emit_op(Op::NULL, line);
        lget(chunk, slot, line);
        call_import(chunks, current, "vybe:database", "query", 3, line);
    } else {
        call_import(chunks, current, "vybe:database", "query", 1, line);
    }
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    lget(chunk, stmt_slot, line);
    lget(chunk, rows_slot, line);
    struct_set_key(chunk, "__rows", line);
    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);
    chunk.emit_op(Op::TRUE, line);
    let done = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(not_query);
    lget(chunk, stmt_slot, line);
    if let Some(slot) = params_slot {
        chunk.emit_op(Op::NULL, line);
        lget(chunk, slot, line);
        call_import(chunks, current, "vybe:database", "execute", 3, line);
    } else {
        call_import(chunks, current, "vybe:database", "execute", 1, line);
    }
    let count_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, count_slot, line);
    lget(chunk, stmt_slot, line);
    emit_empty_array(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_set_key(chunk, "__rows", line);
    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);
    lget(chunk, stmt_slot, line);
    lget(chunk, count_slot, line);
    struct_set_key(chunk, "__row_count", line);
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::DYN_NE, line);

    chunk.patch_jump(done);
}

pub fn emit_php_pdo_statement_fetch(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = if argc >= 1 { Some(alloc_local(chunk)) } else { None };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = mode_slot { lset(chunk, slot, line); }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__cursor", line);
    let cursor_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, cursor_slot, line);

    lget(chunk, rows_slot, line);
    lget(chunk, cursor_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, row_slot, line);

    lget(chunk, stmt_slot, line);
    lget(chunk, cursor_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_key(chunk, "__cursor", line);

    if let Some(slot) = mode_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_COLUMN), line);
        chunk.emit_op(Op::DYN_EQ, line);
        let not_column = chunk.emit_jump(Op::BR_IF_FALSE, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        let row_is_null = chunk.emit_jump(Op::BR_IF_TRUE, line);
        emit_first_column_value(chunks, current, row_slot, line);
        let done = chunks[current].emit_jump(Op::BR, line);
        let chunk = &mut chunks[current];
        chunk.patch_jump(row_is_null);
        chunk.emit_op(Op::NULL, line);
        chunk.patch_jump(done);
        chunk.patch_jump(not_column);
    }

    let chunk = &mut chunks[current];
    lget(chunk, row_slot, line);
}

pub fn emit_php_pdo_statement_fetch_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = if argc >= 1 { Some(alloc_local(chunk)) } else { None };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = mode_slot { lset(chunk, slot, line); }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    if let Some(slot) = mode_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_COLUMN), line);
        chunk.emit_op(Op::DYN_EQ, line);
        let not_column = chunk.emit_jump(Op::BR_IF_FALSE, line);

        emit_empty_array(chunks, current, line);
        let out_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);

        push_const(chunk, Value::F64(0.0), line);
        let index_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, index_slot, line);

        lget(chunk, rows_slot, line);
        collections::emit_len(chunks, current, line);
        let len_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);

        let loop_top = chunk.current_offset();
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        chunk.emit_op(Op::DYN_LT, line);
        let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        let skip_push = chunk.emit_jump(Op::BR_IF_TRUE, line);
        lget(chunk, out_slot, line);
        emit_first_column_value(chunks, current, row_slot, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        let chunk = &mut chunks[current];
        chunk.patch_jump(skip_push);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
        chunk.emit_loop(loop_top, line);
        chunk.patch_jump(exit);
        lget(chunk, out_slot, line);
        let done = chunk.emit_jump(Op::BR, line);

        chunk.patch_jump(not_column);
        lget(chunk, rows_slot, line);
        chunk.patch_jump(done);
        return;
    }

    let chunk = &mut chunks[current];
    lget(chunk, rows_slot, line);
}