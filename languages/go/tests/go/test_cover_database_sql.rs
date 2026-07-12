//! database/sql — Open, Query, Exec, Stmt, Tx, Rows breadth compile smokes.

go_compile_cases! {
    // database/sql — package functions
    sql_open_dsn => "package main; import \"database/sql\"; func main() { _, _ = sql.Open(\"stub\", \":memory:\") }",
    sql_drivers => "package main; import \"database/sql\"; func main() { _ = sql.Drivers() }",
    sql_level_repeatable_read => "package main; import \"database/sql\"; func main() { _ = sql.LevelRepeatableRead; _ = sql.LevelWriteCommitted }",

    // database/sql — DB connection management
    sql_db_ping => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.Ping() }",
    sql_db_ping_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.PingContext(context.Background()) }",
    sql_db_close => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); _ = db.Close() }",
    sql_db_set_max_idle_conns => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); db.SetMaxIdleConns(4) }",
    sql_db_set_max_open_conns => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); db.SetMaxOpenConns(8) }",
    sql_db_set_conn_max_lifetime => "package main; import \"database/sql\"; import \"time\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); db.SetConnMaxLifetime(time.Hour) }",
    sql_db_set_conn_max_idle_time => "package main; import \"database/sql\"; import \"time\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); db.SetConnMaxIdleTime(time.Minute) }",
    sql_db_stats => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.Stats() }",
    sql_db_driver => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.Driver() }",

    // database/sql — DB Prepare
    sql_db_prepare => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.Prepare(\"SELECT 1\") }",
    sql_db_prepare_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.PrepareContext(context.Background(), \"SELECT 1\") }",

    // database/sql — DB Exec
    sql_db_exec => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.Exec(\"INSERT INTO t VALUES (1)\") }",
    sql_db_exec_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.ExecContext(context.Background(), \"DELETE FROM t WHERE id = ?\", 1) }",

    // database/sql — DB Query
    sql_db_query => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.Query(\"SELECT id FROM t\") }",
    sql_db_query_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.QueryContext(context.Background(), \"SELECT name FROM t WHERE id = ?\", 2) }",

    // database/sql — DB QueryRow
    sql_db_query_row => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.QueryRow(\"SELECT count(*) FROM t\") }",
    sql_db_query_row_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.QueryRowContext(context.Background(), \"SELECT max(id) FROM t\") }",

    // database/sql — DB transactions
    sql_db_begin => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.Begin() }",
    sql_db_begin_tx => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.BeginTx(context.Background(), &sql.TxOptions{Isolation: sql.LevelSerializable}) }",
    sql_db_conn => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _, _ = db.Conn(context.Background()) }",

    // database/sql — Stmt
    sql_stmt_close => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"SELECT 1\"); if stmt != nil { _ = stmt.Close() } }",
    sql_stmt_exec => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"INSERT INTO t VALUES (?)\"); if stmt != nil { _, _ = stmt.Exec(1) } }",
    sql_stmt_exec_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"UPDATE t SET v = ?\"); if stmt != nil { _, _ = stmt.ExecContext(context.Background(), 2) } }",
    sql_stmt_query => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"SELECT id FROM t WHERE id = ?\"); if stmt != nil { _, _ = stmt.Query(3) } }",
    sql_stmt_query_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"SELECT name FROM t\"); if stmt != nil { _, _ = stmt.QueryContext(context.Background()) } }",
    sql_stmt_query_row => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"SELECT count(*) FROM t\"); if stmt != nil { _ = stmt.QueryRow(5) } }",
    sql_stmt_query_row_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"SELECT min(id) FROM t\"); if stmt != nil { _ = stmt.QueryRowContext(context.Background()) } }",

    // database/sql — Tx
    sql_tx_commit => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _ = tx.Commit() } }",
    sql_tx_rollback => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _ = tx.Rollback() } }",
    sql_tx_prepare => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _, _ = tx.Prepare(\"UPDATE t SET v = ? WHERE id = ?\") } }",
    sql_tx_prepare_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _, _ = tx.PrepareContext(context.Background(), \"UPDATE t SET v = ?\") } }",
    sql_tx_stmt => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); stmt, _ := db.Prepare(\"SELECT 1\"); tx, _ := db.Begin(); if tx != nil && stmt != nil { _ = tx.Stmt(stmt) } }",
    sql_tx_exec => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _, _ = tx.Exec(\"INSERT INTO t VALUES (?)\", 7) } }",
    sql_tx_exec_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _, _ = tx.ExecContext(context.Background(), \"DELETE FROM t WHERE id = ?\", 8) } }",
    sql_tx_query => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _, _ = tx.Query(\"SELECT id FROM t WHERE id > ?\", 0) } }",
    sql_tx_query_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _, _ = tx.QueryContext(context.Background(), \"SELECT name FROM t\") } }",
    sql_tx_query_row => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _ = tx.QueryRow(\"SELECT count(*) FROM t WHERE active = ?\", true) } }",
    sql_tx_query_row_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); tx, _ := db.Begin(); if tx != nil { _ = tx.QueryRowContext(context.Background(), \"SELECT min(id) FROM t\") } }",

    // database/sql — Rows
    sql_rows_close => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { _ = rows.Close() } }",
    sql_rows_columns => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { _, _ = rows.Columns() } }",
    sql_rows_column_types => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { _, _ = rows.ColumnTypes() } }",
    sql_rows_next => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { _ = rows.Next() } }",
    sql_rows_scan => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { var id int; _ = rows.Scan(&id) } }",
    sql_rows_err => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { _ = rows.Err() } }",

    // database/sql — Row
    sql_row_scan => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); row := db.QueryRow(\"SELECT 1\"); var n int; _ = row.Scan(&n) }",
    sql_row_err => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); _ = db.QueryRow(\"SELECT 1\").Err() }",

    // database/sql — Result
    sql_result_last_insert_id => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); res, _ := db.Exec(\"INSERT INTO t VALUES (1)\"); if res != nil { _, _ = res.LastInsertId() } }",
    sql_result_rows_affected => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); res, _ := db.Exec(\"DELETE FROM t\"); if res != nil { _, _ = res.RowsAffected() } }",

    // database/sql — Null types
    sql_null_string => "package main; import \"database/sql\"; type NullString = sql.NullString; func main() { var ns NullString; _ = ns.Valid; _ = ns.String }",
    sql_null_int64 => "package main; import \"database/sql\"; type NullInt64 = sql.NullInt64; func main() { var ni NullInt64; _ = ni.Valid; _ = ni.Int64 }",
    sql_null_float64 => "package main; import \"database/sql\"; type NullFloat64 = sql.NullFloat64; func main() { var nf NullFloat64; _ = nf.Valid; _ = nf.Float64 }",
    sql_null_bool => "package main; import \"database/sql\"; type NullBool = sql.NullBool; func main() { var nb NullBool; _ = nb.Valid; _ = nb.Bool }",
    sql_null_time => "package main; import \"database/sql\"; import \"time\"; type NullTime = sql.NullTime; func main() { var nt NullTime; _ = nt.Valid; _ = nt.Time; _ = time.Now() }",
    sql_null_int32 => "package main; import \"database/sql\"; type NullInt32 = sql.NullInt32; func main() { var ni NullInt32; _ = ni.Valid; _ = ni.Int32 }",

    // database/sql — TxOptions and isolation
    sql_tx_options_isolation => "package main; import \"database/sql\"; type TxOptions = sql.TxOptions; func main() { opts := TxOptions{Isolation: sql.LevelReadCommitted, ReadOnly: true}; _ = opts.Isolation; _ = opts.ReadOnly }",
    sql_isolation_level_string => "package main; import \"database/sql\"; func main() { _ = sql.LevelDefault.String(); _ = sql.LevelReadUncommitted.String() }",

    // database/sql — ColumnType and NamedArg
    sql_column_type_name => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); rows, _ := db.Query(\"SELECT 1\"); if rows != nil { types, _ := rows.ColumnTypes(); if len(types) > 0 { _, _ = types[0].Name() } } }",
    sql_named_arg => "package main; import \"database/sql\"; type NamedArg = sql.NamedArg; func main() { _ = sql.Named(\"id\", 1); _ = NamedArg{Name: \"name\", Value: \"go\"} }",

    // database/sql — Out and RawBytes
    sql_out => "package main; import \"database/sql\"; type Out = sql.Out; func main() { var dest int; _ = Out{Dest: &dest} }",
    sql_raw_bytes => "package main; import \"database/sql\"; type RawBytes = sql.RawBytes; func main() { var rb RawBytes; _ = rb }",

    // database/sql — DBStats fields
    sql_db_stats_fields => "package main; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); s := db.Stats(); _ = s.MaxOpenConnections; _ = s.OpenConnections; _ = s.InUse; _ = s.Idle }",

    // database/sql — Conn methods
    sql_conn_close => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); conn, _ := db.Conn(context.Background()); if conn != nil { _ = conn.Close() } }",
    sql_conn_exec_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); conn, _ := db.Conn(context.Background()); if conn != nil { _, _ = conn.ExecContext(context.Background(), \"UPDATE t SET v = 1\") } }",
    sql_conn_query_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); conn, _ := db.Conn(context.Background()); if conn != nil { _, _ = conn.QueryContext(context.Background(), \"SELECT 1\") } }",
    sql_conn_query_row_context => "package main; import \"context\"; import \"database/sql\"; func main() { db, _ := sql.Open(\"stub\", \":memory:\"); defer db.Close(); conn, _ := db.Conn(context.Background()); if conn != nil { _ = conn.QueryRowContext(context.Background(), \"SELECT 2\") } }",
}
