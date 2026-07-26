use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

#[test]
fn php_autoloading_runtime() {
    for n in 1..=20_i64 {
        assert_int(
            &format!(
                "spl_autoload_register(function (string $class) {{\n    if ($class === 'Autoloaded{n}') {{\n        eval('class Autoloaded{n} {{ public function id(): int {{ return {n}; }} }}');\n    }}\n}});\n$instance = new Autoloaded{n}();\necho $instance->id();"
            ),
            n,
        );
    }
}

#[test]
fn php_autoloading_multiple_loaders() {
    assert_eq!(
        run_prints(
            r#"<?php
$trace = [];
spl_autoload_register(function (string $class) use (&$trace): void {
    if ($class === 'Autoload\\A') {
        $trace[] = 'primary';
        eval('class Autoload\\A {}');
    }
}, true, false);
spl_autoload_register(function (string $class) use (&$trace): void {
    if ($class === 'Autoload\\A') {
        $trace[] = 'secondary';
        eval('class Autoload\\A2 {}');
    }
}, true, false);
$a = new Autoload\\A();
echo implode('|', $trace);
echo $a::class === 'Autoload\\\\A' ? 'ok' : 'bad';
"#
        ),
        vec!["primaryok"]
    );
}

#[test]
fn php_autoloading_no_autoload_when_flag_false() {
    assert_eq!(
        run_prints(
            r#"<?php
spl_autoload_register(function (string $class): void {
    if ($class === 'Autoload\\Missing') {
        eval('class Missing {}');
    }
});
echo class_exists('Autoload\\\\Missing', false) ? 'found' : 'not_found';
echo class_exists('Autoload\\\\Missing', true) ? 'loaded' : 'not_loaded';
"#
        ),
        vec!["not_foundloaded"]
    );
}

#[test]
fn php_autoloading_class_alias() {
    assert_eq!(
        run_prints(
            r#"<?php
class AutoloadBase {}
class_alias(AutoloadBase::class, 'AutoloadAlias');
echo (new AutoloadAlias()) instanceof AutoloadBase ? 'yes' : 'no';
echo class_alias('AutoloadBase', 'AutoloadAlias', false) ? 'second' : 'first_fail';
"#
        ),
        vec!["yesfirst_fail"]
    );
}

#[test]
fn php_autoloading_unregister() {
    assert_eq!(
        run_prints(
            r#"<?php
$trace = [];
$loader = function (string $class) use (&$trace): void {
    $trace[] = $class;
    if ($class === 'Autoload\\Removable') {
        eval('class Autoload\\\\Removable {}');
    }
};
spl_autoload_register($loader);
echo class_exists('Autoload\\Removable', true) ? 'loaded' : 'not';
spl_autoload_unregister($loader);
echo class_exists('Autoload\\Never', true) ? 'bad' : 'missing';
echo implode('|', $trace);
"#
        ),
        vec!["loadedmissing|Autoload\\\\Removable"]
    );
}

#[test]
fn php_autoloading_prepend_affects_loading_order() {
    assert_eq!(
        run_prints(
            r#"<?php
function autoload_order_default(string $class): void {
    if ($class === 'Autoload\\OrderProbe') {
        eval('class Autoload\\\\OrderProbe {}');
    }
}
function autoload_order_prepend(string $class): void {
    if ($class === 'Autoload\\OrderProbe') {
        eval('class Autoload\\\\OrderProbe {}');
    }
}
spl_autoload_register('autoload_order_default');
spl_autoload_register('autoload_order_prepend', true, true);
$functions = spl_autoload_functions();
if (is_array($functions) && count($functions) >= 2) {
    echo (is_array($functions[0]) ? $functions[0][0] : 'none') . '|';
    echo (is_array($functions[1]) ? $functions[1][0] : 'none');
} else {
    echo 'bad';
}
"#
        ),
        vec!["autoload_order_prepend|autoload_order_default"]
    );
}

#[test]
fn php_autoloading_class_exists_second_arg_toggle() {
    assert_eq!(
        run_prints(
            r#"<?php
$called = 0;
spl_autoload_register(function (string $class) use (&$called): void {
    if ($class === 'Autoload\\Maybe') {
        $called++;
        eval('class Autoload\\\\Maybe {}');
    }
});
echo class_exists('Autoload\\Maybe', false) ? 'pre' : 'pre-no';
echo $called . '|';
echo class_exists('Autoload\\Maybe', true) ? 'loaded' : 'noload';
echo $called . '|';
echo class_exists('Autoload\\Maybe', true) ? 'cached' : 'no-cache';
echo $called;
"#
        ),
        vec!["pre-no1|loaded1|cached1"]
    );
}

#[test]
fn php_autoload_functions_list_contains_loader() {
    assert_eq!(
        run_prints(
            r#"<?php
$loader = function (string $class): void {
    if ($class === 'Autoload\\ListMe') { eval('class Autoload\\\\ListMe {}'); }
};
spl_autoload_register($loader);
$functions = spl_autoload_functions();
$found = false;
foreach ($functions as $f) {
    if (is_array($f) && $f[1] === '__invoke') {
        $found = true;
    }
}
echo $found ? 'found' : 'missing';
echo count($functions) >= 1 ? 'yes' : 'no';
"#
        ),
        vec!["foundyes"]
    );
}
