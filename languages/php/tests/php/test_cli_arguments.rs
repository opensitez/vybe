//! CLI script argument handling: `$argc`, `$argv`, and common parsing patterns.

crate::php_cases! {
    argc_matches_argv_count => {
        r#"<?php
$argv = ['script.php', 'run', '--flag'];
$argc = count($argv);
echo $argc;
"#,
        ["3"]
    };

    argv_zero_is_script_name => {
        r#"<?php
$argv = ['mytool.php', 'sub'];
echo $argv[0];
"#,
        ["mytool.php"]
    };

    argv_first_user_argument => {
        r#"<?php
$argv = ['app.php', 'hello', 'world'];
echo $argv[1] . ':' . $argv[2];
"#,
        ["hello:world"]
    };

    argv_missing_index_is_null => {
        r#"<?php
$argv = ['only.php'];
echo isset($argv[1]) ? 'set' : 'unset';
"#,
        ["unset"]
    };

    foreach_argv_collects_flags => {
        r#"<?php
$argv = ['tool.php', '-v', '-q', 'file.txt'];
$flags = '';
foreach ($argv as $i => $arg) {
    if ($i > 0 && str_starts_with($arg, '-')) {
        $flags .= $arg;
    }
}
echo $flags;
"#,
        ["-v-q"]
    };

    array_shift_skips_script_name => {
        r#"<?php
$argv = ['run.php', 'build', 'release'];
array_shift($argv);
echo implode(',', $argv);
"#,
        ["build,release"]
    };

    argc_zero_when_argv_empty => {
        r#"<?php
$argv = [];
$argc = count($argv);
echo $argc;
"#,
        ["0"]
    };

    argv_numeric_strings_stay_strings => {
        r#"<?php
$argv = ['x.php', '42'];
echo gettype($argv[1]) . ':' . $argv[1];
"#,
        ["string:42"]
    };

    parse_argv_style_option_value => {
        r#"<?php
$argv = ['opts.php', '--name=vybe'];
$opt = $argv[1];
$eq = strpos($opt, '=');
echo substr($opt, $eq + 1);
"#,
        ["vybe"]
    };

    argc_gt_one_enables_subcommand => {
        r#"<?php
$argv = ['cli.php', 'migrate'];
echo ($argc = count($argv)) > 1 ? $argv[1] : 'help';
"#,
        ["migrate"]
    };

    argv_slice_user_args => {
        r#"<?php
$argv = ['main.php', 'a', 'b', 'c'];
$user = array_slice($argv, 1);
echo count($user) . ':' . $user[0];
"#,
        ["3:a"]
    };

    argv_last_positional_argument => {
        r#"<?php
$argv = ['deploy.php', 'staging', 'v2'];
echo $argv[count($argv) - 1];
"#,
        ["v2"]
    };

    argv_in_implode_usage_message => {
        r#"<?php
$argv = ['usage.php'];
echo 'Usage: ' . $argv[0] . ' <file>';
"#,
        ["Usage: usage.php <file>"]
    };

    argv_empty_string_argument => {
        r#"<?php
$argv = ['x.php', ''];
echo $argv[1] === '' ? 'empty' : 'nonempty';
"#,
        ["empty"]
    };

    argv_count_after_shift_loop => {
        r#"<?php
$argv = ['p.php', 'one', 'two'];
while (count($argv) > 1) {
    array_shift($argv);
}
echo count($argv);
"#,
        ["1"]
    };

    argv_compare_strict_option => {
        r#"<?php
$argv = ['t.php', '--help'];
echo $argv[1] === '--help' ? 'help' : 'run';
"#,
        ["help"]
    };

    argv_numeric_index_two => {
        r#"<?php
$argv = ['s.php', 'first', 'second'];
echo $argv[2];
"#,
        ["second"]
    };

    argc_equals_argv_length => {
        r#"<?php
$argv = ['a.php', 'b', 'c', 'd'];
echo count($argv) === 4 ? 'match' : 'mismatch';
"#,
        ["match"]
    };

    argv_list_destructure_first_two => {
        r#"<?php
$argv = ['run.php', 'cmd', 'arg'];
[$script, $command] = $argv;
echo $command;
"#,
        ["cmd"]
    };

    argv_filter_non_dash_args => {
        r#"<?php
$argv = ['f.php', 'in.txt', '-o', 'out.txt'];
$files = [];
foreach (array_slice($argv, 1) as $a) {
    if (!str_starts_with($a, '-')) {
        $files[] = $a;
    }
}
echo implode('+', $files);
"#,
        ["in.txt+out.txt"]
    };

    argv_default_when_missing => {
        r#"<?php
$argv = ['d.php'];
$out = $argv[2] ?? 'default.out';
echo $out;
"#,
        ["default.out"]
    };

    argv_join_remaining_for_shell => {
        r#"<?php
$argv = ['wrap.php', 'git', 'status', '--short'];
$cmd = implode(' ', array_slice($argv, 1));
echo $cmd;
"#,
        ["git status --short"]
    };

    argv_bool_flag_detection => {
        r#"<?php
$argv = ['v.php', '-v'];
$verbose = in_array('-v', $argv, true);
echo $verbose ? 'verbose' : 'quiet';
"#,
        ["verbose"]
    };
}
