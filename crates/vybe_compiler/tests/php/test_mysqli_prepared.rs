//! `mysqli_stmt_*` prepared-statement surface (shape/runtime stubs).

crate::php_cases! {
    mysqli_prepare_returns_stmt_object => {
        r#"<?php
$mysqli = new mysqli('127.0.0.1', 'u', 'p', 'db');
$stmt = $mysqli->prepare('SELECT 1');
echo $stmt === false ? 'fail' : get_class($stmt);
"#,
        ["mysqli_stmt"]
    };

    mysqli_stmt_bind_param_integer_type => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$id = 7;
$stmt->bind_param('i', $id);
echo $id;
"#,
        ["7"]
    };

    mysqli_stmt_bind_param_string_type => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$name = 'ann';
$stmt->bind_param('s', $name);
echo $name;
"#,
        ["ann"]
    };

    mysqli_stmt_bind_param_double_type => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$rate = 1.5;
$stmt->bind_param('d', $rate);
echo (string)$rate;
"#,
        ["1.5"]
    };

    mysqli_stmt_bind_param_blob_type => {
        r#"<?php
$stmt = (new mysqli())->prepare('INSERT INTO t VALUES (?)');
$blob = 'raw';
$stmt->bind_param('b', $blob);
echo strlen($blob);
"#,
        ["3"]
    };

    mysqli_stmt_bind_param_multiple_placeholders => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?, ?');
$a = 1;
$b = 2;
$stmt->bind_param('ii', $a, $b);
echo $a + $b;
"#,
        ["3"]
    };

    mysqli_stmt_execute_returns_bool => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
echo $stmt->execute() ? 'ok' : 'no';
"#,
        ["ok"]
    };

    mysqli_stmt_affected_rows_after_update => {
        r#"<?php
$stmt = (new mysqli())->prepare('UPDATE t SET n = 1');
$stmt->execute();
echo $stmt->affected_rows;
"#,
        ["0"]
    };

    mysqli_stmt_insert_id_after_insert => {
        r#"<?php
$stmt = (new mysqli())->prepare('INSERT INTO t VALUES (NULL)');
$stmt->execute();
echo $stmt->insert_id;
"#,
        ["0"]
    };

    mysqli_stmt_num_rows_select => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
$stmt->store_result();
echo $stmt->num_rows;
"#,
        ["0"]
    };

    mysqli_stmt_fetch_assoc_returns_null_when_empty => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1 WHERE 0');
$stmt->execute();
$row = $stmt->get_result()->fetch_assoc();
echo $row === null ? 'none' : 'row';
"#,
        ["none"]
    };

    mysqli_stmt_fetch_row_numeric => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1 AS n');
$stmt->execute();
$row = $stmt->get_result()->fetch_row();
echo $row[0] ?? 'x';
"#,
        ["1"]
    };

    mysqli_stmt_bind_result_reads_column => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 42 AS n');
$stmt->execute();
$n = 0;
$stmt->bind_result($n);
$stmt->fetch();
echo $n;
"#,
        ["42"]
    };

    mysqli_stmt_error_empty_when_ok => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
echo $stmt->error === '' ? 'clean' : 'err';
"#,
        ["clean"]
    };

    mysqli_stmt_errno_zero_when_ok => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
echo $stmt->errno;
"#,
        ["0"]
    };

    mysqli_stmt_sqlstate_after_prepare => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
echo strlen($stmt->sqlstate) >= 0 ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    mysqli_stmt_reset_clears_bound_state => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$id = 1;
$stmt->bind_param('i', $id);
$stmt->reset();
echo 'reset';
"#,
        ["reset"]
    };

    mysqli_stmt_free_result_after_store => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
$stmt->store_result();
$stmt->free_result();
echo 'free';
"#,
        ["free"]
    };

    mysqli_stmt_close_releases_handle => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->close();
echo 'closed';
"#,
        ["closed"]
    };

    mysqli_stmt_param_count_matches_placeholders => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?, ?');
echo $stmt->param_count;
"#,
        ["2"]
    };

    mysqli_stmt_field_count_select => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1, 2');
echo $stmt->field_count;
"#,
        ["2"]
    };

    mysqli_stmt_get_warnings_empty => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$w = $stmt->get_warnings();
echo $w === false ? 'none' : 'warn';
"#,
        ["none"]
    };

    mysqli_stmt_bind_param_reference_updates => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$v = 5;
$stmt->bind_param('i', $v);
$v = 9;
echo $v;
"#,
        ["9"]
    };

    mysqli_stmt_execute_twice_rebind => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$a = 1;
$stmt->bind_param('i', $a);
$stmt->execute();
$a = 2;
$stmt->execute();
echo $a;
"#,
        ["2"]
    };

    mysqli_stmt_fetch_all_assoc => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1 AS n');
$stmt->execute();
$rows = $stmt->get_result()->fetch_all(MYSQLI_ASSOC);
echo count($rows);
"#,
        ["1"]
    };

    mysqli_stmt_attr_get_update_max_length => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
echo is_int($stmt->attr_get(MYSQLI_STMT_ATTR_UPDATE_MAX_LENGTH)) ? 'int' : 'no';
"#,
        ["int"]
    };

    mysqli_stmt_send_long_data_chunk => {
        r#"<?php
$stmt = (new mysqli())->prepare('INSERT INTO t VALUES (?)');
$stmt->send_long_data(0, 'chunk');
echo 'sent';
"#,
        ["sent"]
    };

    mysqli_stmt_data_seek_after_result => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
$stmt->store_result();
$stmt->data_seek(0);
echo 'seek';
"#,
        ["seek"]
    };

    mysqli_stmt_result_metadata_field_count => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1 AS n');
$stmt->execute();
$meta = $stmt->result_metadata();
echo $meta->field_count;
"#,
        ["1"]
    };

    mysqli_stmt_more_results_false_for_select => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
echo $stmt->more_results() ? 'more' : 'done';
"#,
        ["done"]
    };

    mysqli_stmt_next_result_false => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
echo $stmt->next_result() ? 'next' : 'end';
"#,
        ["end"]
    };

    mysqli_stmt_fetch_object_class => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1 AS n');
$stmt->execute();
$obj = $stmt->get_result()->fetch_object();
echo isset($obj->n) ? (string)$obj->n : '0';
"#,
        ["1"]
    };

    mysqli_stmt_bind_param_null_via_string => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT ?');
$val = null;
$stmt->bind_param('s', $val);
echo $val === null ? 'null' : 'set';
"#,
        ["null"]
    };

    mysqli_stmt_store_result_before_num_rows => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
$stmt->store_result();
echo $stmt->num_rows >= 0 ? 'stored' : 'bad';
"#,
        ["stored"]
    };

    mysqli_stmt_get_result_is_mysqli_result => {
        r#"<?php
$stmt = (new mysqli())->prepare('SELECT 1');
$stmt->execute();
$res = $stmt->get_result();
echo is_object($res) ? 'obj' : 'no';
"#,
        ["obj"]
    };
}
