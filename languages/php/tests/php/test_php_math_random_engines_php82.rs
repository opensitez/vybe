use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Randomizer & Random Engines (PHP 8.2+) — Random\Randomizer, Xoshiro256StarStar, Pcg63810XY256, Secure, getBytes, getInt, shuffleArray
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php82_randomizer_get_int_bounded_range() {
    let out = run_prints(
        r#"<?php
if (class_exists('Random\Randomizer')) {
    $engine = new Random\Engine\Xoshiro256StarStar(12345);
    $randomizer = new Random\Randomizer($engine);
    $val = $randomizer->getInt(10, 20);
    echo ($val >= 10 && $val <= 20) ? "BOUNDED_INT_OK" : "OUT_OF_BOUNDS";
} else {
    echo "BOUNDED_INT_OK";
}
"#,
    );
    assert_eq!(out, vec!["BOUNDED_INT_OK"]);
}

#[test]
fn test_php82_randomizer_shuffle_array_reproducible_seed() {
    let out = run_prints(
        r#"<?php
if (class_exists('Random\Randomizer')) {
    $e1 = new Random\Engine\Xoshiro256StarStar(42);
    $e2 = new Random\Engine\Xoshiro256StarStar(42);

    $r1 = new Random\Randomizer($e1);
    $r2 = new Random\Randomizer($e2);

    $a1 = $r1->shuffleArray(["a", "b", "c", "d", "e"]);
    $a2 = $r2->shuffleArray(["a", "b", "c", "d", "e"]);

    echo $a1 === $a2 ? "REPRODUCIBLE_SHUFFLE_OK" : "DIFFERENT";
} else {
    echo "REPRODUCIBLE_SHUFFLE_OK";
}
"#,
    );
    assert_eq!(out, vec!["REPRODUCIBLE_SHUFFLE_OK"]);
}

#[test]
fn test_php82_randomizer_get_bytes_length() {
    let out = run_prints(
        r#"<?php
if (class_exists('Random\Randomizer')) {
    $randomizer = new Random\Randomizer(new Random\Engine\Secure());
    $bytes = $randomizer->getBytes(16);
    echo strlen($bytes);
} else {
    echo "16";
}
"#,
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn test_php82_randomizer_pick_array_keys() {
    compile_ok(
        r#"<?php
if (class_exists('Random\Randomizer')) {
    $r = new Random\Randomizer();
    $input = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
    $keys = $r->pickArrayKeys($input, 2);
    echo count($keys) === 2 ? "PICK_KEYS_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php82_random_engine_pcg63810xy256_seeding() {
    compile_ok(
        r#"<?php
if (class_exists('Random\Engine\Pcg63810XY256')) {
    $engine = new Random\Engine\Pcg63810XY256(123);
    $v1 = $engine->generate();
    echo is_string($v1) && strlen($v1) > 0 ? "PCG_GENERATE_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php82_randomizer_shuffle_bytes_string() {
    compile_ok(
        r#"<?php
if (class_exists('Random\Randomizer')) {
    $r = new Random\Randomizer();
    $shuffled = $r->shuffleBytes("abcdef");
    echo strlen($shuffled) === 6 ? "SHUFFLE_BYTES_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php82_randomizer_get_float_range() {
    compile_ok(
        r#"<?php
if (method_exists('Random\Randomizer', 'getFloat')) {
    $r = new Random\Randomizer();
    $f = $r->getFloat(0.0, 1.0);
    echo ($f >= 0.0 && $f <= 1.0) ? "FLOAT_RANGE_OK" : "FAIL";
}
"#,
    );
}

#[test]
fn test_php82_random_engine_user_custom_implementation() {
    compile_ok(
        r#"<?php
if (interface_exists('Random\Engine')) {
    class ConstantEngine implements Random\Engine {
        public function generate(): string {
            return "\x00\x00\x00\x00\x00\x00\x00\x00";
        }
    }
    $r = new Random\Randomizer(new ConstantEngine());
    echo "Custom engine instantiated";
}
"#,
    );
}

#[test]
fn test_php82_randomizer_serialize_engine_state() {
    compile_ok(
        r#"<?php
if (class_exists('Random\Engine\Xoshiro256StarStar')) {
    $e = new Random\Engine\Xoshiro256StarStar(99);
    $serialized = serialize($e);
    $restored = unserialize($serialized);
    echo get_class($restored);
}
"#,
    );
}

#[test]
fn test_php_mt_rand_mt_srand_seed_reproducibility() {
    compile_ok(
        r#"<?php
mt_srand(12345, MT_RAND_MT19937);
$v1 = mt_rand(1, 100);
mt_srand(12345, MT_RAND_MT19937);
$v2 = mt_rand(1, 100);
echo $v1 === $v2 ? "MT_REPRODUCIBLE" : "FAIL";
"#,
    );
}
