//! PDO prepared statements — surfaces not covered in `test_database.rs` (`bindValue`, `bindColumn`, fetch loops, stmt metadata).

crate::php_cases! {
    prepared_bindvalue_positional_inserts_integer => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE n (v INTEGER)');
$stmt = $pdo->prepare('INSERT INTO n (v) VALUES (?)');
$stmt->bindValue(1, 42, PDO::PARAM_INT);
$stmt->execute();
echo $pdo->query('SELECT v FROM n')->fetchColumn();
"#,
        ["42"]
    };

    prepared_bindvalue_named_inserts_string => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE u (name TEXT)');
$stmt = $pdo->prepare('INSERT INTO u (name) VALUES (:name)');
$stmt->bindValue(':name', 'ada', PDO::PARAM_STR);
$stmt->execute();
echo $pdo->query('SELECT name FROM u')->fetchColumn();
"#,
        ["ada"]
    };

    prepared_bindvalue_param_bool_stores_truthy => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE f (on INTEGER)');
$stmt = $pdo->prepare('INSERT INTO f (on) VALUES (?)');
$stmt->bindValue(1, true, PDO::PARAM_BOOL);
$stmt->execute();
echo $pdo->query('SELECT on FROM f')->fetchColumn();
"#,
        ["1"]
    };

    prepared_bindvalue_param_null_explicit => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE t (v TEXT)');
$stmt = $pdo->prepare('INSERT INTO t (v) VALUES (?)');
$stmt->bindValue(1, null, PDO::PARAM_NULL);
$stmt->execute();
$row = $pdo->query('SELECT v FROM t')->fetch(PDO::FETCH_ASSOC);
echo $row['v'] === null ? 'null' : 'set';
"#,
        ["null"]
    };

    prepared_bindparam_reads_variable_by_reference_at_execute => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE r (id INTEGER, label TEXT)');
$pdo->exec("INSERT INTO r VALUES (1, 'one'), (2, 'two')");
$stmt = $pdo->prepare('SELECT label FROM r WHERE id = ?');
$id = 1;
$stmt->bindParam(1, $id, PDO::PARAM_INT);
$id = 2;
$stmt->execute();
echo $stmt->fetchColumn();
"#,
        ["two"]
    };

    prepared_bindvalue_snapshots_value_not_reference => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE r (id INTEGER, label TEXT)');
$pdo->exec("INSERT INTO r VALUES (1, 'one'), (2, 'two')");
$stmt = $pdo->prepare('SELECT label FROM r WHERE id = ?');
$id = 1;
$stmt->bindValue(1, $id, PDO::PARAM_INT);
$id = 2;
$stmt->execute();
echo $stmt->fetchColumn();
"#,
        ["one"]
    };

    prepared_while_fetch_accumulates_labels => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE x (n TEXT)');
$pdo->exec("INSERT INTO x VALUES ('a'), ('b')");
$stmt = $pdo->prepare('SELECT n FROM x ORDER BY n');
$stmt->execute();
$out = '';
while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) { $out .= $row['n']; }
echo $out;
"#,
        ["ab"]
    };

    prepared_fetchall_num_indexed_rows => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE p (a INTEGER, b INTEGER)');
$pdo->exec('INSERT INTO p VALUES (1, 10), (2, 20)');
$stmt = $pdo->prepare('SELECT a, b FROM p ORDER BY a');
$stmt->execute();
$rows = $stmt->fetchAll(PDO::FETCH_NUM);
echo $rows[1][1];
"#,
        ["20"]
    };

    prepared_fetch_key_pair_builds_map => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE kv (k TEXT, v TEXT)');
$pdo->exec("INSERT INTO kv VALUES ('x', '1'), ('y', '2')");
$stmt = $pdo->prepare('SELECT k, v FROM kv');
$stmt->execute();
$map = $stmt->fetchAll(PDO::FETCH_KEY_PAIR);
echo $map['y'];
"#,
        ["2"]
    };

    prepared_column_count_matches_select_list => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE w (a TEXT, b TEXT, c TEXT)');
$stmt = $pdo->prepare('SELECT a, b, c FROM w');
$stmt->execute();
echo $stmt->columnCount();
"#,
        ["3"]
    };

    prepared_param_count_matches_placeholders => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$stmt = $pdo->prepare('SELECT ? AS one, ? AS two');
echo $stmt->paramCount();
"#,
        ["2"]
    };

    prepared_delete_row_count_reports_removed => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE d (id INTEGER)');
$pdo->exec('INSERT INTO d VALUES (1), (2), (3)');
$stmt = $pdo->prepare('DELETE FROM d WHERE id > ?');
$stmt->execute([1]);
echo $stmt->rowCount();
"#,
        ["2"]
    };

    prepared_update_zero_rows_row_count_zero => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE u (id INTEGER, v TEXT)');
$pdo->exec("INSERT INTO u VALUES (1, 'a')");
$stmt = $pdo->prepare("UPDATE u SET v = 'b' WHERE id = ?");
$stmt->execute([99]);
echo $stmt->rowCount();
"#,
        ["0"]
    };

    prepared_limit_offset_placeholders => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE s (id INTEGER)');
$pdo->exec('INSERT INTO s VALUES (1), (2), (3), (4)');
$stmt = $pdo->prepare('SELECT id FROM s ORDER BY id LIMIT ? OFFSET ?');
$stmt->execute([2, 1]);
echo implode(',', $stmt->fetchAll(PDO::FETCH_COLUMN));
"#,
        ["2,3"]
    };

    prepared_reexecute_same_statement_different_params => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE q (code TEXT)');
$stmt = $pdo->prepare('INSERT INTO q (code) VALUES (?)');
$stmt->execute(['A']);
$stmt->execute(['B']);
$stmt->execute(['C']);
echo $pdo->query('SELECT COUNT(*) FROM q')->fetchColumn();
"#,
        ["3"]
    };

    prepared_select_no_rows_fetch_returns_false => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE e (id INTEGER)');
$stmt = $pdo->prepare('SELECT id FROM e WHERE id = ?');
$stmt->execute([1]);
echo $stmt->fetch() === false ? 'false' : 'row';
"#,
        ["false"]
    };

    prepared_execute_no_placeholders_empty_array => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE c (v INTEGER)');
$pdo->exec('INSERT INTO c VALUES (9)');
$stmt = $pdo->prepare('SELECT v FROM c');
$stmt->execute([]);
echo $stmt->fetchColumn();
"#,
        ["9"]
    };

    prepared_insert_within_transaction_visible_after_commit => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE t (v TEXT)');
$ins = $pdo->prepare('INSERT INTO t (v) VALUES (?)');
$pdo->beginTransaction();
$ins->execute(['held']);
$pdo->commit();
echo $pdo->query('SELECT v FROM t')->fetchColumn();
"#,
        ["held"]
    };

    prepared_fetch_column_from_statement_not_query => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE m (n INTEGER)');
$pdo->exec('INSERT INTO m VALUES (4), (5)');
$stmt = $pdo->prepare('SELECT SUM(n) FROM m');
$stmt->execute();
echo $stmt->fetchColumn();
"#,
        ["9"]
    };

    prepared_bindcolumn_fills_php_variable_on_fetch => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE b (name TEXT)');
$pdo->exec("INSERT INTO b VALUES ('bound')");
$stmt = $pdo->prepare('SELECT name FROM b');
$stmt->bindColumn(1, $col);
$stmt->execute();
$stmt->fetch(PDO::FETCH_BOUND);
echo $col;
"#,
        ["bound"]
    };

    prepared_error_code_zero_after_successful_execute => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE z (v INTEGER)');
$stmt = $pdo->prepare('INSERT INTO z (v) VALUES (?)');
$stmt->execute([1]);
echo $stmt->errorCode();
"#,
        ["00000"]
    };

    prepared_execute_constraint_violation_sets_error_info => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_SILENT);
$pdo->exec('CREATE TABLE u (id INTEGER PRIMARY KEY)');
$pdo->exec('INSERT INTO u (id) VALUES (1)');
$stmt = $pdo->prepare('INSERT INTO u (id) VALUES (?)');
$stmt->execute([1]);
$info = $stmt->errorInfo();
echo ($info[0] ?? '') !== '00000' ? 'err' : 'ok';
"#,
        ["err"]
    };

    prepared_attr_emulate_prepares_still_executes_insert => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->setAttribute(PDO::ATTR_EMULATE_PREPARES, true);
$pdo->exec('CREATE TABLE e (t TEXT)');
$stmt = $pdo->prepare('INSERT INTO e (t) VALUES (?)');
$stmt->execute(['emu']);
echo $pdo->query('SELECT t FROM e')->fetchColumn();
"#,
        ["emu"]
    };

    prepared_last_insert_id_after_prepared_insert => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE ai (id INTEGER PRIMARY KEY AUTOINCREMENT, v TEXT)');
$stmt = $pdo->prepare('INSERT INTO ai (v) VALUES (?)');
$stmt->execute(['first']);
echo $pdo->lastInsertId();
"#,
        ["1"]
    };

    prepared_mixed_named_params_in_execute_array => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE pair (a TEXT, b TEXT)');
$stmt = $pdo->prepare('INSERT INTO pair (a, b) VALUES (:first, :second)');
$stmt->execute([':first' => 'X', ':second' => 'Y']);
$row = $pdo->query('SELECT a, b FROM pair')->fetch(PDO::FETCH_NUM);
echo $row[0] . $row[1];
"#,
        ["XY"]
    };

    prepared_select_for_update_style_row_lock_not_required_sqlite => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE bal (id INTEGER, amt INTEGER)');
$pdo->exec('INSERT INTO bal VALUES (1, 100)');
$sel = $pdo->prepare('SELECT amt FROM bal WHERE id = ?');
$sel->execute([1]);
$upd = $pdo->prepare('UPDATE bal SET amt = amt - ? WHERE id = ?');
$upd->execute([30, 1]);
$sel->execute([1]);
echo $sel->fetchColumn();
"#,
        ["70"]
    };

    mysqli_prepare_returns_stmt_object_shape => {
        r#"<?php
$dbh = mysqli_init();
$stmt = mysqli_prepare($dbh, 'SELECT 1');
echo is_object($stmt) ? 'stmt' : 'no';
"#,
        ["stmt"]
    };
}
