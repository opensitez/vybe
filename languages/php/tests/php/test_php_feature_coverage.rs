use super::helpers::run_prints;

const PHP_FEATURE_VALUES: [i64; 8] = [1, 2, 3, 7, 16, 42, 60, 120];

fn assert_output(expr: &str, expected: &str) {
    assert_eq!(
        run_prints(&format!("<?php echo {}; ", expr)),
        vec![expected.to_string()]
    );
}

fn assert_int(expr: &str, expected: i64) {
    assert_output(expr, &expected.to_string());
}

fn assert_str(expr: &str, expected: &str) {
    assert_output(expr, expected);
}

fn assert_bool(expr: &str, expected: bool) {
    assert_output(expr, if expected { "1" } else { "0" });
}

fn decimal_len(mut value: i64) -> i64 {
    if value == 0 {
        return 1;
    }
    if value < 0 {
        value = -value;
    }
    value.to_string().len() as i64
}

#[test]
fn php_arithmetic_operators() {
    for n in PHP_FEATURE_VALUES {
        assert_int(&format!("({n} + 3) * 2 - 1"), n * 2 + 5);
        assert_int(&format!("{n} * 4 - 7"), n * 4 - 7);
        assert_int(&format!("intdiv({n} * 100, 25)"), n * 4);
        assert_int(&format!("({n} + 8) % 7"), (n + 8) % 7);
    }
}

#[test]
fn php_bitwise_operators() {
    for n in PHP_FEATURE_VALUES {
        assert_int(&format!("({n} ^ 3) & 7"), (n ^ 3) & 7);
        assert_int(&format!("({n} << 1) >> 1"), (n << 1) >> 1);
        assert_int(
            &format!("({n} | 1) - ({n} & 1)"),
            if n % 2 == 0 { n } else { n } + 1 - 1,
        );
        assert_int(&format!("(~{n}) & 31"), (!n) & 31);
    }
}

#[test]
fn php_comparison_operators() {
    for n in PHP_FEATURE_VALUES {
        assert_bool(&format!("({n} % 2 === 0)"), n % 2 == 0);
        assert_bool(&format!("({n} > 100)"), n > 100);
        let cmp_result = match n.cmp(&60) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        };
        assert_int(&format!("({n} <=> 60)"), cmp_result);
        assert_bool(&format!("({n} != 42)"), n != 42);
        assert_int(&format!("({n} <=> 60)"), cmp_result);
    }
}

#[test]
fn php_arrays_data_structures() {
    for n in PHP_FEATURE_VALUES {
        let len = 3 + (n % 5);
        let range_end = 1 + (n % 5);
        let index = n % 2;

        assert_int(&format!("count(array_fill(0, {len}, {n}))"), len);
        assert_int(
            &format!("array_sum(range(1, {range_end}))"),
            (range_end * (range_end + 1)) / 2,
        );
        assert_int(
            &format!("$arr = [0 => {n}, 1 => {m}]; echo count($arr);", m = n + 1),
            2,
        );
        assert_int(
            &format!(
                "$arr = [0, 2, 4]; $arr[{index}] = $arr[{index}] + {n} + 1; echo $arr[{index}];"
            ),
            n + 1
                + if index == 0 {
                    0
                } else if index == 1 {
                    2
                } else {
                    4
                },
        );
    }
}

#[test]
fn php_string_features() {
    for n in PHP_FEATURE_VALUES {
        assert_int(&format!("strlen('x{n}')"), decimal_len(n) + 1);
        assert_int(
            &format!("strlen(str_replace('x', '', 'x{n}x'))"),
            decimal_len(n),
        );
        assert_str(&format!("strtoupper(substr('ab{n}', 0, 2))"), "AB");
        assert_int(
            &format!("strlen(implode('-', ['a', '{n}', 'c']))"),
            1 + decimal_len(n) + 1 + 1 + 1,
        );
        assert_int(&format!("count(explode(',', 'a,{n},c'))"), 3);
        assert_str(&format!("substr('prefix-{n}-suffix', 0, 6)"), "prefix");
        assert_str(&format!("str_repeat('*', 2)"), "**");
    }
}

#[test]
fn php_for_loop_constructs() {
    for n in PHP_FEATURE_VALUES {
        let expected = n * (n - 1) / 2;
        assert_int(
            &format!(
                " $total = 0; for($i = 0; $i < {n}; $i++) : $total += $i; endfor; echo $total;"
            ),
            expected,
        );
    }
}

#[test]
fn php_foreach_loop_constructs() {
    for n in PHP_FEATURE_VALUES {
        let steps = n % 7 + 1;
        let expected = steps * n;
        assert_int(
            &format!(
                "$items = array_fill(0, {steps}, {n});\n$sum = 0;\nforeach ($items as $item) :\n    $sum += $item;\nendforeach;\necho $sum;"
            ),
            expected,
        );
    }
}

#[test]
fn php_scope_rules() {
    for n in PHP_FEATURE_VALUES {
        let expected = if n % 3 == 0 { n + 10 } else { n + 1 };
        assert_int(
            &format!(
                "$value = {n};\nif ($value % 3 === 0) :\n    $value += 10;\nelse :\n    $value += 1;\nendif;\necho $value;"
            ),
            expected,
        );
        assert_int(
            &format!("$value = {n}; if ($value > 100) {{ $value = 100; }} echo $value;"),
            if n > 100 { 100 } else { n },
        );
    }
}

#[test]
fn php_oop_classes() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!(
                "class SurfaceCounter{{ public function __construct(private int $seed) {{}} public function value(): int {{ return $this->seed; }} }}\n$instance = new SurfaceCounter({n}); echo $instance->value();"
            ),
            n,
        );
    }
}

#[test]
fn php_namespaces() {
    for n in PHP_FEATURE_VALUES {
        assert_str(
            &format!(
                "namespace N{n}; echo (int)\\DateTime::createFromFormat('U', '0')->format('Y');"
            ),
            "1970",
        );
    }
}

#[test]
fn php_dynamic_calling() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!("$fn = 'strlen'; echo $fn('x{n}');"),
            decimal_len(n) + 1,
        );
        assert_int(
            &format!(
                "$obj = new DateTime('@'.({n} + 2000)); $method = 'getTimestamp'; echo $obj->$method() - 2000;"
            ),
            n,
        );
        assert_str(
            &format!(
                "call_user_func(['DateTime', 'createFromFormat'], 'U', '{n}') !== null ? 'ok' : 'no';"
            ),
            "ok",
        );
    }
}

#[test]
fn php_output_buffering() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!("ob_start(); echo {n}; $out = ob_get_clean(); echo strlen($out);"),
            decimal_len(n),
        );
    }
}

#[test]
fn php_method_chaining() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!(
                "$dt = (new DateTime('@'.({n} + 3000)))->setTimezone(new DateTimeZone('UTC'));\necho $dt->getTimestamp() - 3000;"
            ),
            n,
        );
    }
}

#[test]
fn php_spl_data_structures() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!(
                "$q = new SplQueue(); $q->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_KEEP); $q->enqueue({n}); echo $q->count();"
            ),
            1,
        );
        assert_int(
            &format!(
                "$stack = new SplStack(); $stack->push({n}); $stack->push({n} + 1); echo $stack->pop();"
            ),
            n + 1,
        );
        assert_int(
            &format!(
                "$heap = new SplPriorityQueue(); $heap->setExtractFlags(SplPriorityQueue::EXTR_BOTH); $heap->insert({n}, -{n}); $top = $heap->extract(); echo $top['data'];"
            ),
            n,
        );
        assert_int(
            &format!(
                "$obj = new SplObjectStorage(); $a = new DateTime('@{n}'); $obj[$a] = 'ok'; echo $obj->count();"
            ),
            1,
        );
    }
}

#[test]
fn php_fiber_async_runtime() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!(
                "$f = new Fiber(function(int $value): int {{\n    return $value + 7;\n}});\n$f->start({n});\necho $f->getReturn();"
            ),
            n + 7,
        );
    }
}

#[test]
fn php_math_builtins() {
    for n in PHP_FEATURE_VALUES {
        assert_int(&format!("abs(-{n}) + 3"), n + 3);
        assert_int(&format!("pow({n}, 2) / {n}"), n);
        assert_int(&format!("min({n}, {n} + 3)"), n);
        assert_int(&format!("max({n} - 2, {n})"), n);
    }
}

#[test]
fn php_datetime_features() {
    for n in PHP_FEATURE_VALUES {
        assert_int(
            &format!(
                "$ts = 1_700_000_000 + {n}; $dt = new DateTime('@' . $ts); echo $dt->getTimestamp() - 1_700_000_000;"
            ),
            n,
        );
        assert_str(
            &format!("$tz = new DateTimeZone('UTC'); echo $tz->getName();"),
            "UTC",
        );
        let idx = n % 3;
        assert_int(
            &format!(
                "$zones = ['UTC', 'Europe/London', 'Asia/Tokyo']; echo strlen($zones[{idx}]);"
            ),
            match idx {
                0 => 3,
                1 => 12,
                _ => 8,
            },
        );
        assert_int(
            &format!(
                "$dt = new DateTime('1970-01-01 00:00:00', new DateTimeZone('UTC')); echo (int)$dt->format('U');"
            ),
            0,
        );
    }
}

#[test]
fn php_literals_and_casting() {
    for n in PHP_FEATURE_VALUES {
        assert_int(&format!("(int)'{n}'"), n);
        assert_int(&format!("(int)(((float)'{n}.5') - ({n}.0))"), 0);
        assert_bool(&format!("({n} > 0 && '{n}' == \"{n}\")"), true);
        let key = if n % 2 == 0 { "'bool'" } else { "'null'" };
        assert_str(
            &format!("['null' => null, 'bool' => true][{key}]"),
            if n % 2 == 0 { "1" } else { "" },
        );
    }
}

#[test]
fn php_timezone_features() {
    for n in PHP_FEATURE_VALUES {
        assert_int(&format!("echo {n};"), n);
        assert_str(
            &format!(
                "$tz = new DateTimeZone('America/New_York'); echo substr($tz->getName(), 0, 3);"
            ),
            "Ame",
        );
        assert_int(
            &format!(
                "$base = new DateTime('2024-01-01 12:00:00', new DateTimeZone('Europe/London')); $base->modify('+{n} days'); $local = new DateTimeZone('Asia/Tokyo'); echo (string)$base->setTimezone($local)->format('Y');"
            ),
            2024,
        );
        assert_int(
            &format!(
                "$tz = new DateTimeZone('UTC'); echo (int)$tz->getOffset(new DateTime('1970-01-01 00:00:00', $tz));"
            ),
            0,
        );
    }
}
